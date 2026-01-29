import io
import re
import tempfile
from typing import List, Dict, Tuple

import pandas as pd
import streamlit as st
import pdfplumber


# -------------------------
# Texto / normalização
# -------------------------
def _normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


# -------------------------
# Índice (página 1) -> páginas corretas
# -------------------------
def _pages_from_index(pdf_path: str) -> Dict[str, List[int]]:
    """
    Tenta ler a página 1 (ÍNDICE) e extrair números de página.
    Retorna {"balanco":[ativo, passivo], "dre":[x], "dfc":[y]} quando encontrar.
    """
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            return {"balanco": [], "dre": [], "dfc": []}

        txt = _normalize_text(pdf.pages[0].extract_text() or "")

    # Ex.: "balanço patrimonial ativo 2"
    m_ba = re.search(r"balan[çc]o patrimonial ativo\s+(\d+)", txt)
    m_bp = re.search(r"balan[çc]o patrimonial passivo\s+(\d+)", txt)
    m_dre = re.search(r"demonstra[çc][ãa]o do resultado\s+(\d+)", txt)

    # Ex.: "demonstração do fluxo de caixa (método indireto) 8"
    m_dfc = re.search(r"demonstra[çc][ãa]o do fluxo de caixa.*?\s(\d+)", txt)

    pages = {"balanco": [], "dre": [], "dfc": []}

    if m_ba:
        pages["balanco"].append(int(m_ba.group(1)))
    if m_bp:
        pages["balanco"].append(int(m_bp.group(1)))
    if m_dre:
        pages["dre"].append(int(m_dre.group(1)))
    if m_dfc:
        pages["dfc"].append(int(m_dfc.group(1)))

    # remove duplicatas/ordena
    for k in pages:
        pages[k] = sorted(list(dict.fromkeys(pages[k])))

    return pages


def _fallback_find_pages(pdf_path: str) -> Dict[str, List[int]]:
    """
    Fallback: se não conseguir ler o índice, faz busca simples,
    mas pega apenas a PRIMEIRA ocorrência real (evita pegar dezenas de páginas).
    """
    pages = {"balanco": [], "dre": [], "dfc": []}

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            txt = _normalize_text(page.extract_text() or "")

            # Evita index (pág 1) e coisas muito genéricas
            if i == 1:
                continue

            if not pages["balanco"] and ("balanço patrimonial" in txt or "balancos patrimonial" in txt):
                pages["balanco"].append(i)
                pages["balanco"].append(i + 1)  # geralmente ativo/passivo em sequência

            if not pages["dre"] and ("demonstração do resultado" in txt or "demonstrações dos resultados" in txt):
                pages["dre"].append(i)

            if not pages["dfc"] and ("fluxo de caixa" in txt or "fluxos de caixa" in txt):
                pages["dfc"].append(i)

    for k in pages:
        pages[k] = sorted(list(dict.fromkeys(pages[k])))

    return pages


def _find_statement_pages(pdf_path: str) -> Dict[str, List[int]]:
    # 1) tenta pelo índice (mais confiável)
    pages = _pages_from_index(pdf_path)

    # se vier vazio demais, usa fallback
    if (not pages["balanco"] and not pages["dre"] and not pages["dfc"]):
        pages = _fallback_find_pages(pdf_path)

    return pages


# -------------------------
# Limpeza robusta
# -------------------------
def _br_number_to_float(x):
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

    # BR: milhares '.' e decimal ','
    s2 = s.replace(".", "").replace(",", ".")
    if not re.fullmatch(r"-?\d+(\.\d+)?", s2):
        return x

    v = float(s2)
    if neg:
        v = -v
    if abs(v - round(v)) < 1e-9:
        return int(round(v))
    return v


def _clean_table_robust(df: pd.DataFrame) -> pd.DataFrame:
    """
    100% robusto:
    - não promove cabeçalho (evita colunas duplicadas)
    - força colunas C0..Cn
    - converte números br
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    if df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = [f"C{i}" for i in range(df.shape[1])]

    for c in df.columns:
        df[c] = df[c].apply(_br_number_to_float)

    return df.reset_index(drop=True)


# -------------------------
# Extração com Camelot
# -------------------------
def _extract_tables_camelot(pdf_path: str, pages: List[int]) -> List[pd.DataFrame]:
    if not pages:
        return []

    import camelot

    page_str = ",".join(map(str, pages))
    dfs: List[pd.DataFrame] = []

    # lattice -> stream fallback
    for flavor in ("lattice", "stream"):
        try:
            tables = camelot.read_pdf(pdf_path, pages=page_str, flavor=flavor)
            for t in tables:
                d = _clean_table_robust(t.df)
                if not d.empty:
                    dfs.append(d)
            if dfs:
                break
        except Exception:
            continue

    return dfs


# -------------------------
# Processamento multi-PDF -> Excel
# -------------------------
def process_multiple_pdfs(files) -> Tuple[bytes, pd.DataFrame]:
    out = io.BytesIO()
    status_rows = []

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # garante pelo menos uma aba
        pd.DataFrame({"info": ["Gerado pelo extrator (Balanço/DRE/DFC)."]}).to_excel(
            writer, sheet_name="INFO", index=False
        )

        for file in files:
            try:
                file.seek(0)
                pdf_bytes = file.read()
                if not pdf_bytes:
                    raise ValueError("Arquivo veio vazio (0 bytes).")

                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    tmp.write(pdf_bytes)
                    pdf_path = tmp.name

                pages = _find_statement_pages(pdf_path)

                base = re.sub(r"[^A-Za-z0-9_]+", "_", file.name.replace(".pdf", ""))[:18] or "arquivo"

                for key, label in [("balanco", "BAL"), ("dre", "DRE"), ("dfc", "DFC")]:
                    pgs = pages.get(key, [])
                    tables = _extract_tables_camelot(pdf_path, pgs)

                    sheet = f"{base}_{label}"[:31]

                    if not tables:
                        pd.DataFrame(
                            {"info": [f"Nenhuma tabela encontrada ({label}).", f"Páginas: {pgs}"]}
                        ).to_excel(writer, sheet_name=sheet, index=False)
                        status_rows.append(
                            {"arquivo": file.name, "demonstracao": label, "status": "ok_sem_tabela", "paginas": str(pgs)}
                        )
                    else:
                        # concat seguro (todas têm C0..Cn)
                        df_all = pd.concat(tables, ignore_index=True)
                        df_all.insert(0, "_arquivo", file.name)
                        df_all.insert(1, "_demo", label)
                        df_all.insert(2, "_paginas", str(pgs))

                        df_all.to_excel(writer, sheet_name=sheet, index=False)
                        status_rows.append(
                            {"arquivo": file.name, "demonstracao": label, "status": "ok", "paginas": str(pgs)}
                        )

            except Exception as e:
                err_sheet = f"ERRO_{re.sub(r'[^A-Za-z0-9_]+','_', file.name)[:24]}"[:31]
                pd.DataFrame({"arquivo": [file.name], "erro": [repr(e)]}).to_excel(
                    writer, sheet_name=err_sheet, index=False
                )
                status_rows.append(
                    {"arquivo": file.name, "demonstracao": "-", "status": "erro", "paginas": "", "erro": repr(e)}
                )

        status_df = pd.DataFrame(status_rows) if status_rows else pd.DataFrame(
            [{"arquivo": "-", "demonstracao": "-", "status": "sem_resultados"}]
        )
        status_df.to_excel(writer, sheet_name="STATUS", index=False)

    return out.getvalue(), status_df


# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="Extrator PDF → Excel (DRE/Balanço/DFC)", layout="wide")
st.title("Extrair DRE, Balanço e DFC de PDFs e gerar Excel")

uploaded_files = st.file_uploader(
    "Selecione os PDFs (pode ser múltiplo)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} arquivo(s) selecionado(s).")

    if st.button("Processar todos"):
        with st.spinner("Extraindo e montando o Excel..."):
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
