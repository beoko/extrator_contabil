import io
import re
import tempfile
from typing import List, Dict, Tuple

import pandas as pd
import streamlit as st
import pdfplumber


# =========================
# Utils
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
# DETECÇÃO DE PÁGINAS (CORRIGIDA)
# =========================
BAL_KEYS = ["balanços patrimoniais", "balancos patrimoniais", "balanço patrimonial", "balanco patrimonial"]
DRE_KEYS = ["demonstrações dos resultados", "demonstracoes dos resultados", "demonstração do resultado", "demonstracao do resultado"]
DFC_KEYS = [
    "demonstrações dos fluxos de caixa", "demonstracoes dos fluxos de caixa",
    "demonstração dos fluxos de caixa", "demonstracao dos fluxos de caixa",
    "demonstração do fluxo de caixa", "demonstracao do fluxo de caixa",
]

INDEX_MARKERS = ["conteúdo", "conteudo", "índice", "indice", "composição do capital", "composicao do capital"]


def is_index_like(t: str) -> bool:
    """
    Heurística: página de sumário/índice tem marcadores + muitos números pequenos de página (2,4,6...).
    """
    tt = norm(t)
    if any(m in tt for m in INDEX_MARKERS):
        # se tem vários números de 1-2 dígitos, típico de sumário
        small_nums = re.findall(r"\b\d{1,2}\b", tt)
        if len(small_nums) >= 8:
            return True
        return True
    return False


def count_numeric_tokens(t: str) -> int:
    """
    Conta "números contábeis" (ex: 17.546, 167.871, 1.234.567, 123.456)
    """
    tt = norm(t)
    # milhares com ponto (pt-BR) ou números grandes
    nums = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d{4,}\b", tt)
    return len(nums)


def score_page_for_statement(t: str, keys: List[str]) -> int:
    """
    Pontua páginas que parecem uma demonstração:
    - contém keyword
    - contém '(em milhares' (muito comum nas DFs)
    - tem muitos números (densidade)
    - não é índice
    """
    tt = norm(t)
    if is_index_like(tt):
        return -999

    s = 0
    if any(k in tt for k in keys):
        s += 50
    if "em milhares" in tt:
        s += 30

    # densidade numérica
    n = count_numeric_tokens(tt)
    if n >= 15:
        s += 30
    elif n >= 8:
        s += 15
    elif n >= 4:
        s += 5

    # pistas úteis
    if "controladora" in tt or "consolidado" in tt:
        s += 5
    if "ativo" in tt or "passivo" in tt:
        s += 3

    return s


def find_statement_pages(pdf_path: str) -> Dict[str, List[int]]:
    """
    Seleciona páginas do PDF por evidência de demonstração (não pelo índice).
    Balanço geralmente 2 páginas (ou 1 com ativo+passivo), DFC às vezes 2.
    """
    pages = {"balanco": [], "dre": [], "dfc": []}

    with pdfplumber.open(pdf_path) as pdf:
        texts = [(p.extract_text() or "") for p in pdf.pages]
        n_pages = len(texts)

    # escolhe a melhor página (maior score) para cada demo
    def best_page(keys: List[str]) -> int | None:
        best_i = None
        best_s = -10**9
        for i, t in enumerate(texts, start=1):
            s = score_page_for_statement(t, keys)
            if s > best_s:
                best_s = s
                best_i = i
        # exige um score mínimo decente para não pegar lixo
        return best_i if best_s >= 55 else None

    bal_p = best_page(BAL_KEYS)
    dre_p = best_page(DRE_KEYS)
    dfc_p = best_page(DFC_KEYS)

    # BAL: tenta pegar 2 páginas se a seguinte também parece tabela/continuação
    if bal_p:
        pages["balanco"].append(bal_p)
        if bal_p < n_pages:
            s_next = score_page_for_statement(texts[bal_p], BAL_KEYS)  # texts index = bal_p (próxima página)
            # se a próxima tem muitos números e não é índice, assume continuação (ativo/passivo)
            if not is_index_like(texts[bal_p]) and count_numeric_tokens(texts[bal_p]) >= 8 and s_next >= 20:
                pages["balanco"].append(bal_p + 1)

    # DRE: normalmente 1 página; inclui próxima se ainda estiver forte
    if dre_p:
        pages["dre"].append(dre_p)
        if dre_p < n_pages:
            if count_numeric_tokens(texts[dre_p]) >= 8 and not is_index_like(texts[dre_p]):
                # só inclui se não virou "resultado abrangente"/outra seção
                if "resultado abrangente" not in norm(texts[dre_p]):
                    pages["dre"].append(dre_p + 1)

    # DFC: pode ser 1 ou 2 páginas
    if dfc_p:
        pages["dfc"].append(dfc_p)
        if dfc_p < n_pages:
            if count_numeric_tokens(texts[dfc_p]) >= 8 and not is_index_like(texts[dfc_p]):
                pages["dfc"].append(dfc_p + 1)

    # limpar duplicatas e páginas fora do range
    for k in pages:
        pages[k] = sorted(list(dict.fromkeys([p for p in pages[k] if 1 <= p <= n_pages])))

    return pages


# =========================
# EXTRAÇÃO POR COORDENADAS (pdfplumber)
# =========================
def cluster_columns(xs: List[float], tol: float = 18.0) -> List[float]:
    xs = sorted(xs)
    if not xs:
        return []
    centers = [xs[0]]
    for x in xs[1:]:
        if abs(x - centers[-1]) <= tol:
            centers[-1] = (centers[-1] + x) / 2
        else:
            centers.append(x)
    return centers


def assign_col(x: float, centers: List[float]) -> int:
    if not centers:
        return 0
    return min(range(len(centers)), key=lambda i: abs(x - centers[i]))


def extract_table_xy(pdf_path: str, page_num: int) -> pd.DataFrame:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_num - 1]
        words = page.extract_words(use_text_flow=True, keep_blank_chars=False)

    if not words:
        return pd.DataFrame()

    # filtra topo/rodapé (remove títulos e numeração de página)
    filtered = []
    for w in words:
        if (w.get("text") or "").strip() == "":
            continue
        if w["top"] < 60:  # topo
            continue
        if w["top"] > (page.height - 35):  # rodapé
            continue
        filtered.append(w)

    if not filtered:
        return pd.DataFrame()

    xs = [w["x0"] for w in filtered]
    centers = cluster_columns(xs, tol=18.0)

    rows: Dict[float, List[Tuple[int, str]]] = {}
    for w in filtered:
        y = round(w["top"], 1)
        c = assign_col(w["x0"], centers)
        rows.setdefault(y, []).append((c, w["text"]))

    y_sorted = sorted(rows.keys())
    ncols = max(len(centers), max((max(c for c, _ in rows[y]) + 1) for y in y_sorted))

    matrix = []
    for y in y_sorted:
        line = [""] * ncols
        items = rows[y]
        tmp = {}
        for c, t in items:
            tmp.setdefault(c, []).append(t)
        for c, parts in tmp.items():
            line[c] = " ".join(parts).strip()
        # exige pelo menos 2 células preenchidas para ser "linha de tabela"
        if sum(1 for v in line if str(v).strip()) >= 2:
            matrix.append(line)

    if not matrix:
        return pd.DataFrame()

    df = pd.DataFrame(matrix, columns=[f"C{i}" for i in range(len(matrix[0]))])
    for c in df.columns:
        df[c] = df[c].apply(br_to_number)
    return df


def extract_statement(pdf_path: str, pages: List[int]) -> pd.DataFrame:
    blocks = []
    for p in pages:
        dfp = extract_table_xy(pdf_path, p)
        if dfp.empty:
            continue
        dfp.insert(0, "_page", p)
        blocks.append(dfp)
        blocks.append(pd.DataFrame([[""] * len(dfp.columns)], columns=dfp.columns))  # separador

    if not blocks:
        return pd.DataFrame()

    df = pd.concat(blocks, ignore_index=True)
    df.columns = make_unique_columns(df.columns)
    return df


# =========================
# MULTI-PDF -> EXCEL
# =========================
def process_multiple_pdfs(files) -> Tuple[bytes, pd.DataFrame]:
    out = io.BytesIO()
    status_rows = []

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame({"info": ["Gerado pelo extrator (Balanço/DRE/DFC) - correção anti-sumário."]}).to_excel(
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
# UI
# =========================
st.set_page_config(page_title="Extrator Contábil (PDF → Excel)", layout="wide")
st.title("Extrator Contábil: Balanço, DRE e DFC (PDF → Excel)")

uploaded_files = st.file_uploader(
    "Selecione os PDFs (pode múltiplo)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} arquivo(s) selecionado(s).")

    if st.button("Processar todos"):
        with st.spinner("Processando (ignorando sumário e buscando páginas reais)..."):
            xlsx_bytes, status_df = process_multiple_pdfs(uploaded_files)

        st.success("Concluído!")
        st.dataframe(status_df, use_container_width=True)

        st.download_button(
            "Baixar Excel consolidado",
            data=xlsx_bytes,
            file_name="demonstracoes_consolidadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.caption("Dica: selecione vários PDFs de uma vez na janela de upload.")
