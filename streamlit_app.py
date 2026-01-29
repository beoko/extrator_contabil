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
    """
    Varre o PDF e tenta localizar páginas por palavras-chave.
    Retorna dict: {"balanco":[..], "dre":[..], "dfc":[..]}
    """
    pages = {"balanco": [], "dre": [], "dfc": []}

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            txt = _normalize_text(page.extract_text() or "")

            # Balanço patrimonial (às vezes split em 2 páginas)
            if "balanços patrimoniais" in txt or "balanço patrimonial" in txt:
                pages["balanco"].append(i)

            # DRE
            if "demonstrações dos resultados" in txt or "demonstração do resultado" in txt:
                pages["dre"].append(i)

            # DFC
            if "demonstrações dos fluxos de caixa" in txt or "demonstração dos fluxos de caixa" in txt:
                pages["dfc"].append(i)

    # Se achou 1 página de balanço, geralmente é 2 (ativo/passivo): adiciona a próxima
    if len(pages["balanco"]) == 1:
        pages["balanco"].append(pages["balanco"][0] + 1)

    for k in pages:
        pages[k] = sorted(list(dict.fromkeys(pages[k])))

    return pages


# -------------------------
# Helpers (tabelas / limpeza)
# -------------------------
def _make_unique_columns(cols) -> List[str]:
    """
    Garante nomes de colunas únicos (evita InvalidIndexError ao escrever Excel).
    """
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
    """
    Converte:
      '1.234.567' -> 1234567
      '(1.234)'   -> -1234
      '1.234,56'  -> 1234.56
    Se não for número, retorna original.
    """
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

    # remove espaços internos
    s = re.sub(r"\s+", "", s)

    # separador BR
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
    Limpa e tenta melhorar tabela:
    - remove linhas vazias
    - usa 1a linha como header se parecer cabeçalho (datas etc.)
    - converte números
    - garante colunas únicas
    """
    if df is None or df.empty:
        return pd.DataFrame()

    # remove linhas totalmente vazias
    df = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    if df.empty:
        return pd.DataFrame()

    # normaliza para string
    df = df.astype(str)

    # tenta promover primeira linha como header se tiver datas ou títulos típicos
    first = [str(x).lower() for x in df.iloc[0].tolist()]
    header_hint = any(("31/12" in c) or ("30/06" in c) or ("controladora" in c) or ("consolidado" in c) for c in first)

    if header_hint:
        df.columns = df.iloc[0]
        df = df.iloc[1:].copy()

    # garante colunas únicas sempre
    df.columns = _make_unique_columns(df.columns)

    # converte números “BR” onde fizer sentido
    for c in df.columns:
        df[c] = df[c].apply(_br_number_to_float)

    return df.reset_index(drop=True)


def _extract_tables_camelot(pdf_path: str, pages: List[int]) -> List[pd.DataFrame]:
    """
    Extrai tabelas com Camelot (lattice -> stream fallback).
    Retorna lista de DataFrames já limpos.
    """
    if not pages:
        return []

    # Import aqui para não quebrar import geral caso Camelot não esteja instalado
    import camelot

    page_str = ",".join(map(str, pages))
    dfs: List[pd.DataFrame] = []

    # 1) lattice
    try:
        tables = camelot.read_pdf(pdf_path, pages=page_str, flavor="lattice")
        for t in tables:
            d = _clean_table(t.df)
            if not d.empty:
                dfs.append(d)
    except Exception:
        pass

    # 2) stream
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
# Core (multi PDF -> Excel)
# -------------------------
def process_multiple_pdfs(files) -> Tuple[bytes, pd.DataFrame]:
    """
    Processa N PDFs:
      - Cria um Excel com abas <arquivo>_BAL / _DRE / _DFC
      - Sempre cria aba INFO (evita 'At least one sheet must be visible')
      - Se um PDF falhar, cria aba ERRO_<arquivo>
    Retorna: (xlsx_bytes, status_df)
    """
    out = io.BytesIO()
    status_rows = []

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # Garante 1 aba sempre
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

                # base curta para nome de aba (máx 31 chars no Excel)
                base = re.sub(r"[^A-Za-z0-9_]+", "_", file.name.replace(".pdf", ""))[:18]
                if not base:
                    base = "arquivo"

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
                        df_all = pd.concat(tables, ignore_index=True)

                        # ✅ garante colunas únicas antes de salvar
                        df_all.columns = _make_unique_columns(df_all.columns)

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

        # Aba de status geral
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
    "Faça upload de **vários PDFs** (ex.: 5 arquivos). Vou tentar detectar as páginas de **Balanço**, **DRE** e **DFC** "
    "e extrair as tabelas para um único Excel."
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
    st.caption("Dica: você pode selecionar vários PDFs de uma vez na janela de upload.")
