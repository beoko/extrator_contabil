import io
import re
import tempfile
from typing import List, Dict, Tuple

import pandas as pd
import streamlit as st
import pdfplumber


# -------------------------
# Helpers
# -------------------------
def _normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _make_unique_columns(cols) -> List[str]:
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

    s2 = s.replace(".", "").replace(",", ".")
    if not re.fullmatch(r"-?\d+(\.\d+)?", s2):
        return x

    v = float(s2)
    if neg:
        v = -v
    if abs(v - round(v)) < 1e-9:
        return int(round(v))
    return v


def _clean_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpa e padroniza:
    - remove vazios
    - tenta promover header
    - garante colunas únicas
    - converte números pt-BR
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    if df.empty:
        return pd.DataFrame()

    # Heurística simples de header (datas, controladora, consolidado etc.)
    first = df.iloc[0].astype(str).str.lower().tolist()
    header_hint = any(
        ("31/12" in c) or ("30/06" in c) or ("31/03" in c) or
        ("controladora" in c) or ("consolidado" in c) or ("nota" in c)
        for c in first
    )
    if header_hint:
        df.columns = df.iloc[0]
        df = df.iloc[1:].copy()

    # Colunas únicas SEMPRE (resolve InvalidIndexError)
    df.columns = _make_unique_columns(df.columns)

    for c in df.columns:
        df[c] = df[c].apply(_br_number_to_float)

    return df.reset_index(drop=True)


def _find_statement_pages(pdf_path: str) -> Dict[str, List[int]]:
    """
    Detecção padrão (pros PDFs que já estavam funcionando pra você).
    """
    pages = {"balanco": [], "dre": [], "dfc": []}

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            txt = _normalize_text(page.extract_text() or "")

            if "balanços patrimoniais" in txt or "balanço patrimonial" in txt:
                if not pages["balanco"]:
                    pages["balanco"].append(i)
                    pages["balanco"].append(i + 1)  # geralmente ativo/passivo

            if ("demonstrações dos resultados" in txt or "demonstração do resultado" in txt) and not pages["dre"]:
                pages["dre"].append(i)

            if ("demonstrações dos fluxos de caixa" in txt or "demonstração do fluxo de caixa" in txt) and not pages["dfc"]:
                pages["dfc"].append(i)
                pages["dfc"].append(i + 1)

    # limpa duplicatas e páginas fora do range
    with pdfplumber.open(pdf_path) as pdf:
        n = len(pdf.pages)
    for k in pages:
        pages[k] = sorted(list(dict.fromkeys([p for p in pages[k] if 1 <= p <= n])))

    return pages


def _extract_tables_camelot(pdf_path: str, pages: List[int]) -> List[pd.DataFrame]:
    """
    Camelot lattice -> stream fallback.
    """
    if not pages:
        return []

    import camelot
    page_str = ",".join(map(str, pages))
    dfs: List[pd.DataFrame] = []

    # lattice
    try:
        tables = camelot.read_pdf(pdf_path, pages=page_str, flavor="lattice")
        for t in tables:
            d = _clean_table(t.df)
            if not d.empty:
                dfs.append(d)
    except Exception:
        pass

    # stream fallback
    if not dfs:
        try:
            tables = camelot.read_pdf(pdf_path, pages=page_str, flavor="stream")
            for t in tables:
                d = _clean_table(t.df)
                if not d.empty:
                    dfs.append(d)
        except Exception:
            pass

    return dfs


# -------------------------
# Core multi-PDF
# -------------------------
def process_multiple_pdfs(files) -> bytes:
    out = io.BytesIO()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # garante 1 sheet sempre
        pd.DataFrame({"info": ["Extrator Balanço/DRE/DFC (Camelot)"]}).to_excel(
            writer, sheet_name="INFO", index=False
        )

        for file in files:
            file.seek(0)
            pdf_bytes = file.read()

            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(pdf_bytes)
                pdf_path = tmp.name

            # ✅ PATCH: se for o 6.pdf, força páginas (ajuste se precisar)
            fname = (file.name or "").lower()
            if fname.endswith("6.pdf") or fname == "6.pdf" or "6.pdf" in fname:
                pages = {"balanco": [5, 6], "dre": [7], "dfc": [8, 9]}
            else:
                pages = _find_statement_pages(pdf_path)

            base = re.sub(r"[^A-Za-z0-9_]+", "_", file.name.replace(".pdf", ""))[:18] or "arquivo"

            for key, label in [("balanco", "BAL"), ("dre", "DRE"), ("dfc", "DFC")]:
                tables = _extract_tables_camelot(pdf_path, pages.get(key, []))
                sheet = f"{base}_{label}"[:31]

                if not tables:
                    pd.DataFrame(
                        {"info": [f"Nenhuma tabela encontrada ({label}). Páginas: {pages.get(key, [])}"]}
                    ).to_excel(writer, sheet_name=sheet, index=False)
                    continue

                # ✅ garante colunas únicas antes de concatenar
                for i in range(len(tables)):
                    tables[i].columns = _make_unique_columns(tables[i].columns)

                df_all = pd.concat(tables, ignore_index=True)
                df_all.columns = _make_unique_columns(df_all.columns)

                df_all.to_excel(writer, sheet_name=sheet, index=False)

    return out.getvalue()


# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="Extrator PDF → Excel", layout="wide")
st.title("Extrator PDF → Excel (Balanço / DRE / DFC)")

uploaded_files = st.file_uploader(
    "Selecione os PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} arquivos selecionados")

    if st.button("Processar todos"):
        with st.spinner("Extraindo..."):
            excel_bytes = process_multiple_pdfs(uploaded_files)

        st.success("Concluído!")
        st.download_button(
            "Baixar Excel consolidado",
            data=excel_bytes,
            file_name="demonstracoes_consolidadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
