# streamlit_app.py
import io
import re
import tempfile
from typing import List, Dict, Tuple

import pandas as pd
import streamlit as st
import pdfplumber


# -------------------------
# Helpers (texto / páginas)
# -------------------------
def _normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _find_statement_pages(pdf_path: str) -> Dict[str, List[int]]:
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

    # Balanço costuma ser 2 páginas (ativo e passivo)
    if len(pages["balanco"]) == 1:
        pages["balanco"].append(pages["balanco"][0] + 1)

    for k in pages:
        pages[k] = sorted(list(dict.fromkeys(pages[k])))

    return pages


# -------------------------
# Helpers (limpeza robusta)
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
    Versão 100% robusta contra colunas duplicadas:
    - NÃO promove primeira linha como header (evita nomes duplicados)
    - Força colunas para C0..Cn sempre
    - Converte números BR célula a célula (quando possível)
    """
    if df is None or df.empty:
        return pd.DataFrame()

    # remove linhas totalmente vazias
    df = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    if df.empty:
        return pd.DataFrame()

    # padroniza colunas
    df = df.copy()
    df.columns = [f"C{i}" for i in range(df.shape[1])]

    # converte números (sem assumir tipos)
    for c in df.columns:
        df[c] = df[c].apply(_br_number_to_float)

    return df.reset_index(drop=True)


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
# Core (multi PDF -> Excel)
# -------------------------
def process_multiple_pdfs(files) -> Tuple[bytes, pd.DataFrame]:
    out = io.BytesIO()
    status_rows = []

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # garante pelo menos uma aba
        pd.DataFrame({"info": ["Gerado pelo extrator (Balanço/DRE/DFC)."]}).to_excel(
            writer, sheet_name="INFO", index=False
        )

        if not files:
            pd.DataFrame({"info": ["Nenhum PDF enviado."]}).to_excel(
                writer, sheet_name="SEM_ARQUIVOS", index=False
            )
            return out.getvalue(), pd.DataFrame([{"arquivo": "-", "status": "sem_arquivos", "detalhe": ""}])

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
                    tables = _extract_tables_camelot(pdf_path, pages.get(key, []))
                    sheet = f"{base}_{label}"[:31]

                    if not tables:
                        pd.DataFrame(
                            {
                                "info": [
                                    f"Nenhuma tabela encontrada ({label}).",
                                    f"Páginas detectadas: {pages.get(key, [])}",
                                ]
                            }
                        ).to_excel(writer, sheet_name=sheet, index=False)

                        status_rows.append(
                            {
                                "arquivo": file.name,
                                "demonstracao": label,
                                "status": "ok_sem_tabela",
                                "paginas": str(pages.get(key, [])),
                                "detalhe": "",
                            }
                        )
                    else:
                        # concat agora é seguro porque TODAS têm colunas C0..Cn
                        df_all = pd.concat(tables, ignore_index=True)

                        # adiciona colunas de rastreio (opcional, mas ajuda muito)
                        df_all.insert(0, "_arquivo", file.name)
                        df_all.insert(1, "_demo", label)

                        df_all.to_excel(writer, sheet_name=sheet, index=False)

                        status_rows.append(
                            {
                                "arquivo": file.name,
                                "demonstracao": label,
                                "status": "ok",
                                "paginas": str(pages.get(key, [])),
                                "detalhe": f"{len(tables)} tabela(s)",
                            }
                        )

            except Exception as e:
                err_sheet = f"ERRO_{re.sub(r'[^A-Za-z0-9_]+','_', file.name)[:24]}"[:31]
                pd.DataFrame({"arquivo": [file.name], "erro": [repr(e)]}).to_excel(
                    writer, sheet_name=err_sheet, index=False
                )
                status_rows.append(
                    {"arquivo": file.name, "demonstracao": "-", "status": "erro", "paginas": "", "detalhe": repr(e)}
                )

        status_df = pd.DataFrame(status_rows) if status_rows else pd.DataFrame(
            [{"arquivo": "-", "demonstracao": "-", "status": "sem_resultados", "paginas": "", "detalhe": ""}]
        )
        status_df.to_excel(writer, sheet_name="STATUS", index=False)

    return out.getvalue(), status_df


# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="Extrator PDF → Excel (DRE/Balanço/DFC)", layout="wide")
st.title("Extrair DRE, Balanço e DFC de PDFs e gerar Excel")

st.write(
    "Faça upload de **vários PDFs**. O app detecta páginas de **Balanço**, **DRE** e **DFC** "
    "e tenta extrair as tabelas para um único Excel."
)

uploaded_files = st.file_uploader(
    "Selecione os PDFs",
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
    st.caption("Dica: selecione vários PDFs de uma vez no upload.")
