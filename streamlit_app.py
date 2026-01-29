import io
import re
import pandas as pd
import streamlit as st
import pdfplumber
from typing import List

# -------------------------
# Helpers
# -------------------------
def _normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _find_statement_pages(pdf_path: str) -> dict:
    pages = {"balanco": [], "dre": [], "dfc": []}

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            txt = _normalize_text(page.extract_text() or "")

            if "balanços patrimoniais" in txt or "balanço patrimonial" in txt:
                pages["balanco"].append(i)

            if "demonstrações dos resultados" in txt or "demonstração do resultado" in txt:
                pages["dre"].append(i)

            if "demonstrações dos fluxos de caixa" in txt or "demonstração dos fluxos de caixa" in txt:
                pages["dfc"].append(i)

    if len(pages["balanco"]) == 1:
        pages["balanco"].append(pages["balanco"][0] + 1)

    for k in pages:
        pages[k] = sorted(list(set(pages[k])))

    return pages


def _br_number_to_float(x):
    if x is None:
        return x

    s = str(x).strip()
    if s == "" or s == "-":
        return None

    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]

    s = s.replace("R$", "").replace(".", "").replace(",", ".")

    try:
        v = float(s)
        return -v if neg else v
    except:
        return x


def _clean_table(df: pd.DataFrame) -> pd.DataFrame:
    df = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    if df.empty:
        return df

    first = df.iloc[0].astype(str).str.lower().tolist()
    if any("31/12" in c or "30/06" in c for c in first):
        df.columns = df.iloc[0]
        df = df.iloc[1:]

    for c in df.columns:
        df[c] = df[c].apply(_br_number_to_float)

    return df.reset_index(drop=True)


def _extract_tables(pdf_path: str, pages: List[int]) -> List[pd.DataFrame]:
    if not pages:
        return []

    import camelot
    page_str = ",".join(map(str, pages))
    tables = []

    try:
        tables = camelot.read_pdf(pdf_path, pages=page_str, flavor="lattice")
    except:
        pass

    if not tables:
        try:
            tables = camelot.read_pdf(pdf_path, pages=page_str, flavor="stream")
        except:
            pass

    return [_clean_table(t.df) for t in tables if not t.df.empty]


# -------------------------
# Core multi-PDF
# -------------------------
def process_multiple_pdfs(files) -> bytes:
    import tempfile
    out = io.BytesIO()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for file in files:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(file.read())
                pdf_path = tmp.name

            pages = _find_statement_pages(pdf_path)

            base = file.name.replace(".pdf", "")[:20]

            for key, label in [("balanco", "BAL"), ("dre", "DRE"), ("dfc", "DFC")]:
                tables = _extract_tables(pdf_path, pages[key])

                if not tables:
                    pd.DataFrame(
                        {"info": [f"Nenhuma tabela encontrada ({label})"]}
                    ).to_excel(
                        writer, sheet_name=f"{base}_{label}", index=False
                    )
                    continue

                df_all = pd.concat(tables, ignore_index=True)
                df_all.to_excel(
                    writer,
                    sheet_name=f"{base}_{label}"[:31],
                    index=False
                )

    return out.getvalue()


# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="Extrator PDF → Excel", layout="wide")
st.title("Upload múltiplo de PDFs – DRE, Balanço e DFC")

st.write(
    "Faça upload de **vários PDFs ao mesmo tempo**. "
    "Vou extrair Balanço, DRE e DFC de cada um e gerar **um único Excel**."
)

uploaded_files = st.file_uploader(
    "Selecione os PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} arquivos selecionados")

    if st.button("Processar todos"):
        with st.spinner("Extraindo dados..."):
            excel_bytes = process_multiple_pdfs(uploaded_files)

        st.success("Processamento concluído!")
        st.download_button(
            "Baixar Excel consolidado",
            data=excel_bytes,
            file_name="demonstracoes_consolidadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
