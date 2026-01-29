# streamlit_app.py
import io
import re
import math
import tempfile
from typing import List, Dict, Tuple, Optional

import pandas as pd
import streamlit as st
import pdfplumber


# =========================
# Utilitários gerais
# =========================
def norm(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def safe_sheet(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_]+", "_", name).strip("_")
    return (name or "ARQ")[:31]


def make_unique_columns(cols) -> List[str]:
    seen = {}
    out = []
    for c in list(cols):
        c = "COL" if c is None or str(c).strip() == "" else str(c).strip()
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out


def br_to_number(x):
    """Converte números pt-BR e parênteses negativo. Se não for número, mantém."""
    if x is None:
        return None
    s = str(x).strip()
    if s == "" or s == "-":
        return None

    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()

    s = s.replace("R$", "").strip()
    s = re.sub(r"\s+", "", s)

    s2 = s.replace(".", "").replace(",", ".")
    if not re.fullmatch(r"-?\d+(\.\d+)?", s2):
        return x

    v = float(s2)
    if neg:
        v = -v
    if abs(v - round(v)) < 1e-9:
        return int(round(v))
    return v


# =========================
# 1) Encontrar páginas (índice + fallback)
# =========================
INDEX_KEYS = {
    "balanco": ["balanços patrimoniais", "balancos patrimoniais", "balanço patrimonial", "balanco patrimonial"],
    "dre": ["demonstrações dos resultados", "demonstracoes dos resultados", "demonstração do resultado", "demonstracao do resultado"],
    "dfc": ["demonstrações dos fluxos de caixa", "demonstracoes dos fluxos de caixa", "demonstração dos fluxos de caixa", "demonstracao dos fluxos de caixa", "demonstração do fluxo de caixa", "demonstracao do fluxo de caixa"],
}

# padrões de "próxima seção" pra saber quando parar páginas contínuas
STOP_TITLES = [
    "demonstrações dos resultados abrangentes",
    "demonstracoes dos resultados abrangentes",
    "demonstrações das mutações do patrimônio líquido",
    "demonstracoes das mutacoes do patrimonio liquido",
    "demonstrações do valor adicionado",
    "demonstracoes do valor adicionado",
    "notas explicativas",
]


def pages_from_index(pdf: pdfplumber.PDF) -> Dict[str, List[int]]:
    """
    Lê a página do Índice (geralmente página 2) e tenta capturar o número da página.
    Ex.: "Balanços patrimoniais .......... 7"
    """
    out = {"balanco": [], "dre": [], "dfc": []}

    # tenta achar uma página que tenha "Índice"
    idx_page = None
    for i in range(min(4, len(pdf.pages))):
        t = norm(pdf.pages[i].extract_text() or "")
        if "índice" in t or "indice" in t:
            idx_page = i
            break

    if idx_page is None:
        return out

    txt = norm(pdf.pages[idx_page].extract_text() or "")
    # pega linhas e tenta achar "... <numero>"
    lines = [l.strip() for l in txt.split(" ")]

    # melhor: regex no texto inteiro por cada chave
    for k, keys in INDEX_KEYS.items():
        for kw in keys:
            m = re.search(rf"{re.escape(kw)}.*?(\d{{1,3}})\b", txt)
            if m:
                out[k] = [int(m.group(1))]
                break
        # não continua se já achou
    return out


def page_has_any(text: str, patterns: List[str]) -> bool:
    t = norm(text)
    return any(p in t for p in patterns)


def find_statement_pages(pdf_path: str) -> Dict[str, List[int]]:
    """
    Estratégia:
    1) tenta Índice: pega a página inicial de cada demo (muito confiável em 1.pdf..5.pdf)
    2) expande "contíguas" até detectar outra seção/título
    3) fallback: varrer páginas e pegar primeira ocorrência real dos títulos
    """
    pages = {"balanco": [], "dre": [], "dfc": []}

    with pdfplumber.open(pdf_path) as pdf:
        idx = pages_from_index(pdf)

        # ---------- expandir por contiguidade ----------
        def expand_from(start_page_1based: int, include_keywords: List[str]) -> List[int]:
            if start_page_1based <= 0 or start_page_1based > len(pdf.pages):
                return []
            outp = []
            i = start_page_1based
            while 1 <= i <= len(pdf.pages):
                txt = norm(pdf.pages[i - 1].extract_text() or "")
                # condição de permanência: ainda tem keyword (ou é continuação imediata da mesma tabela)
                # e NÃO entrou em outra seção (STOP)
                has_kw = any(kw in txt for kw in include_keywords)
                has_stop = any(stp in txt for stp in STOP_TITLES)

                # regra prática:
                # - sempre inclui a página inicial
                # - inclui mais 1 página se a inicial tem tabela e a próxima é continuação (às vezes sem repetir título)
                if i == start_page_1based:
                    outp.append(i)
                    i += 1
                    continue

                # se a página seguinte ainda tem tabela/continuação: muitos PDFs repetem o cabeçalho da demo
                if has_kw and not has_stop:
                    outp.append(i)
                    i += 1
                    continue

                # caso específico: balanço frequentemente quebra em 2 páginas (ativo/passivo) ou
                # a continuação não repete o título, mas repete “Controladora / Consolidado / Nota / Ativo / Passivo”.
                cont_hint = ("controladora" in txt or "consolidado" in txt or "ativo" in txt or "passivo" in txt) and not has_stop
                if cont_hint and (i == start_page_1based + 1):
                    outp.append(i)
                    i += 1
                    continue

                break

            return outp

        # se índice achou, usa como base
        if idx["balanco"]:
            # balanço às vezes é 2 páginas de tabela; expand resolve
            pages["balanco"] = expand_from(idx["balanco"][0], INDEX_KEYS["balanco"])
        if idx["dre"]:
            pages["dre"] = expand_from(idx["dre"][0], INDEX_KEYS["dre"])
        if idx["dfc"]:
            pages["dfc"] = expand_from(idx["dfc"][0], INDEX_KEYS["dfc"])

        # ---------- fallback: varrer se faltou ----------
        if not pages["balanco"] or not pages["dre"] or not pages["dfc"]:
            for i, p in enumerate(pdf.pages, start=1):
                txt = norm(p.extract_text() or "")
                if not pages["balanco"] and page_has_any(txt, INDEX_KEYS["balanco"]):
                    pages["balanco"] = [i, min(i + 1, len(pdf.pages))]
                if not pages["dre"] and page_has_any(txt, INDEX_KEYS["dre"]):
                    pages["dre"] = [i]
                if not pages["dfc"] and page_has_any(txt, INDEX_KEYS["dfc"]):
                    pages["dfc"] = [i, min(i + 1, len(pdf.pages))]
                if pages["balanco"] and pages["dre"] and pages["dfc"]:
                    break

    # limpar duplicatas/ordenar
    for k in pages:
        pages[k] = sorted(list(dict.fromkeys([p for p in pages[k] if p > 0])))

    return pages


# =========================
# 2) Extrair tabela por coordenadas (robusto)
# =========================
def cluster_columns(xs: List[float], tol: float = 18.0) -> List[float]:
    """
    Agrupa posições x em "colunas" por proximidade.
    tol em pontos (PDF units). Ajuste se precisar.
    """
    xs = sorted(xs)
    if not xs:
        return []
    centers = [xs[0]]
    for x in xs[1:]:
        if abs(x - centers[-1]) <= tol:
            # atualiza centro por média simples
            centers[-1] = (centers[-1] + x) / 2
        else:
            centers.append(x)
    return centers


def assign_col(x: float, centers: List[float]) -> int:
    if not centers:
        return 0
    j = min(range(len(centers)), key=lambda i: abs(x - centers[i]))
    return j


def extract_table_xy(pdf_path: str, page_num: int) -> pd.DataFrame:
    """
    Extrai "tabela" a partir de palavras e suas coordenadas.
    Funciona melhor que Camelot em PDFs contábeis com colunas múltiplas.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_num - 1]
        words = page.extract_words(use_text_flow=True, keep_blank_chars=False)

    if not words:
        return pd.DataFrame()

    # filtra coisas muito no rodapé (número de página isolado)
    # e muito no topo (título)
    # (ajuste leve pra não jogar fora header de tabela)
    filtered = []
    for w in words:
        if w.get("text", "").strip() == "":
            continue
        if w["top"] < 55:  # topo
            continue
        if w["top"] > (page.height - 35):  # rodapé
            continue
        filtered.append(w)

    if not filtered:
        return pd.DataFrame()

    # 1) detectar colunas por x0 (começo das palavras)
    xs = [w["x0"] for w in filtered]
    centers = cluster_columns(xs, tol=18.0)

    # 2) agrupar linhas por y (top arredondado)
    rows: Dict[float, List[Tuple[int, str]]] = {}
    for w in filtered:
        y = round(w["top"], 1)
        col = assign_col(w["x0"], centers)
        rows.setdefault(y, []).append((col, w["text"]))

    # 3) ordenar linhas e montar matriz
    y_sorted = sorted(rows.keys())
    ncols = max(len(centers), max((max(c for c, _ in rows[y]) + 1) for y in y_sorted))

    matrix = []
    for y in y_sorted:
        line = [""] * ncols
        items = sorted(rows[y], key=lambda x: x[0])
        # concatena textos que caem na mesma coluna
        tmp = {}
        for c, t in items:
            tmp.setdefault(c, []).append(t)
        for c, parts in tmp.items():
            line[c] = " ".join(parts).strip()
        # remove linhas praticamente vazias
        if sum(1 for v in line if str(v).strip()) >= 2:
            matrix.append(line)

    if not matrix:
        return pd.DataFrame()

    df = pd.DataFrame(matrix, columns=[f"C{i}" for i in range(len(matrix[0]))])

    # tenta converter números
    for c in df.columns:
        df[c] = df[c].apply(br_to_number)

    return df


def extract_statement(pdf_path: str, pages: List[int]) -> pd.DataFrame:
    """
    Extrai uma demonstração juntando as páginas:
    - adiciona marcadores de página
    - concatena com linha em branco entre páginas
    """
    blocks = []
    for p in pages:
        dfp = extract_table_xy(pdf_path, p)
        if dfp.empty:
            continue
        dfp.insert(0, "_page", p)
        blocks.append(dfp)
        # separador
        blocks.append(pd.DataFrame([[""] * len(dfp.columns)], columns=dfp.columns))

    if not blocks:
        return pd.DataFrame()

    df = pd.concat(blocks, ignore_index=True)
    df.columns = make_unique_columns(df.columns)
    return df


# =========================
# 3) Multi-PDF -> Excel
# =========================
def process_multiple_pdfs(files) -> Tuple[bytes, pd.DataFrame]:
    out = io.BytesIO()
    status_rows = []

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame({"info": ["Gerado pelo extrator (Balanço/DRE/DFC) via coordenadas (pdfplumber)."]}).to_excel(
            writer, sheet_name="INFO", index=False
        )

        for f in files:
            try:
                f.seek(0)
                pdf_bytes = f.read()
                if not pdf_bytes:
                    raise ValueError("Arquivo vazio (0 bytes).")

                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    tmp.write(pdf_bytes)
                    pdf_path = tmp.name

                pages = find_statement_pages(pdf_path)

                base = re.sub(r"\.pdf$", "", f.name, flags=re.I)
                base = safe_sheet(base)[:18] or "ARQ"

                for key, label in [("balanco", "BAL"), ("dre", "DRE"), ("dfc", "DFC")]:
                    pgs = pages.get(key, [])
                    sheet = safe_sheet(f"{base}_{label}")

                    if not pgs:
                        pd.DataFrame({"info": [f"Não encontrei páginas para {label}."]}).to_excel(
                            writer, sheet_name=sheet, index=False
                        )
                        status_rows.append({"arquivo": f.name, "demo": label, "status": "sem_paginas", "paginas": ""})
                        continue

                    df = extract_statement(pdf_path, pgs)

                    if df.empty:
                        pd.DataFrame({"info": [f"Páginas detectadas: {pgs}", "Não consegui extrair tabela nessas páginas."]}).to_excel(
                            writer, sheet_name=sheet, index=False
                        )
                        status_rows.append({"arquivo": f.name, "demo": label, "status": "sem_tabela", "paginas": str(pgs)})
                        continue

                    # salva
                    df.to_excel(writer, sheet_name=sheet, index=False)
                    status_rows.append({"arquivo": f.name, "demo": label, "status": "ok", "paginas": str(pgs)})

            except Exception as e:
                err_sheet = safe_sheet(f"ERRO_{f.name}")[:31]
                pd.DataFrame({"arquivo": [f.name], "erro": [repr(e)]}).to_excel(writer, sheet_name=err_sheet, index=False)
                status_rows.append({"arquivo": f.name, "demo": "-", "status": "erro", "paginas": "", "erro": repr(e)})

        status_df = pd.DataFrame(status_rows) if status_rows else pd.DataFrame([{"status": "sem_resultados"}])
        status_df.to_excel(writer, sheet_name="STATUS", index=False)

    return out.getvalue(), pd.DataFrame(status_rows)


# =========================
# UI Streamlit
# =========================
st.set_page_config(page_title="Extrator Contábil (PDF → Excel)", layout="wide")
st.title("Extrator Contábil: Balanço, DRE e DFC (PDF → Excel)")

st.write(
    "Upload múltiplo de PDFs e geração de Excel consolidado.\n"
    "Este extrator usa **pdfplumber + coordenadas (x/y)**, mais robusto que Camelot para PDFs contábeis."
)

uploaded_files = st.file_uploader(
    "Selecione os PDFs (pode múltiplo)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} arquivo(s) selecionado(s).")

    if st.button("Processar todos"):
        with st.spinner("Processando..."):
            xlsx_bytes, status_df = process_multiple_pdfs(uploaded_files)

        st.success("Concluído!")
        st.dataframe(status_df, use_container_width=True)

        st.download_button(
            "Baixar Excel consolidado",
            data=xlsx_bytes,
            file_name="demonstracoes_consolidadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.caption("Se algum PDF ficar com colunas “esticadas”, dá pra ajustar o `tol` em `cluster_columns()` (18→14 ou 22).")
