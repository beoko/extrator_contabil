import io
import re
import pandas as pd
import streamlit as st
import pdfplumber
from typing import List, Dict, Optional

# -------------------------
# Helpers
# -------------------------
def _normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _extract_lines(page_text: str) -> List[str]:
    return [ln.strip() for ln in (page_text or "").splitlines() if ln.strip()]

def _find_pdf_page_of_report(pdf) -> Optional[int]:
    """
    Tenta localizar onde começa o relatório (auditor/revisão).
    Isso ajuda a calcular o offset entre numeração "impressa" e página real do PDF.
    """
    needles = [
        "relatório do auditor independente",
        "relatório dos auditores independentes",
        "relatório sobre a revisão",
        "relatório sobre a revisão das informações trimestrais",
    ]
    for i, page in enumerate(pdf.pages, start=1):
        txt = _normalize_text(page.extract_text() or "")
        if any(n in txt for n in needles):
            return i
    return None

def _parse_index_page(lines: List[str]) -> Dict[str, int]:
    """
    Extrai os números finais no Índice/Conteúdo.
    Retorna páginas "impressas" (não as páginas reais do PDF).
    """
    joined = " \n".join(lines).lower()

    def find_last_int(patterns: List[str]) -> Optional[int]:
        for pat in patterns:
            m = re.search(pat, joined, flags=re.IGNORECASE)
            if m:
                return int(m.group(1))
        return None

    # Aceita variações comuns
    p_bal = find_last_int([
        r"balanços\s+patrimoniais\s*\.{0,}\s*(\d+)",
        r"balanço\s+patrimonial\s*\.{0,}\s*(\d+)",
    ])
    p_dre = find_last_int([
        r"demonstrações\s+dos\s+resultados\s*\.{0,}\s*(\d+)",
        r"demonstração\s+do\s+resultado\s*\.{0,}\s*(\d+)",
    ])
    p_dfc = find_last_int([
        r"demonstrações\s+dos\s+fluxos\s+de\s+caixa\s*\.{0,}\s*(\d+)",
        r"demonstração\s+dos\s+fluxos\s+de\s+caixa\s*\.{0,}\s*(\d+)",
        r"fluxos\s+de\s+caixa\s*\.{0,}\s*(\d+)",
    ])
    p_rel = find_last_int([
        r"relatório.*?\s(\d+)\b",
    ])

    out = {}
    if p_bal is not None: out["balanco"] = p_bal
    if p_dre is not None: out["dre"] = p_dre
    if p_dfc is not None: out["dfc"] = p_dfc
    if p_rel is not None: out["relatorio"] = p_rel
    return out

def _find_statement_pages(pdf_path: str) -> dict:
    """
    Estratégia:
      1) Tenta pegar páginas via Índice/Conteúdo (mais confiável nesses PDFs).
      2) Converte para páginas reais do PDF via offset.
      3) Fallback: varredura por palavras-chave (teu método antigo).
    """
    pages = {"balanco": [], "dre": [], "dfc": []}

    with pdfplumber.open(pdf_path) as pdf:
        # --- 1) achar página do índice/conteúdo ---
        index_page_nums = []
        for i, page in enumerate(pdf.pages, start=1):
            txt_raw = page.extract_text() or ""
            txt = _normalize_text(txt_raw)
            if ("índice" in txt) or ("indice" in txt) or ("conteúdo" in txt) or ("conteudo" in txt):
                # evita falso positivo no corpo: precisa ter várias ocorrências de "demonstrações" ou "balanços"
                if ("balan" in txt) and ("demonstra" in txt):
                    index_page_nums.append(i)

        report
