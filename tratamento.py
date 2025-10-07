import re
import pandas as pd
import pdfplumber
from pathlib import Path

# >>> CONFIG <<< --------------------------------------------------------------
PDF_PATH = r"01. Relatório de Conformidade - RPV.pdf"
OUT_CSV  = r"/mnt/data/Comparativo_dos_Valores.csv"
OUT_XLSX = r"/mnt/data/Comparativo_dos_Valores.xlsx"
OUT_CSV_LATEST  = r"/mnt/data/Comparativo_dos_Valores_mais_recente.csv"
OUT_XLSX_LATEST = r"/mnt/data/Comparativo_dos_Valores_mais_recente.xlsx"
# ---------------------------------------------------------------------------

HEADER_PATTERN = r"5\.\s*Comparativo dos Valores"
NEXT_HEADER    = r"\n\s*6\.\s*Observações|\Z"
NUM_PATTERN    = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
DATE_PATTERN   = r"\d{2}/\d{2}/\d{4}"

def extract_section_text(pdf_path: str) -> str:
    texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or "")
    full = "\n".join(texts)
    m = re.search(HEADER_PATTERN + r"(.*?)" + NEXT_HEADER, full, flags=re.S)
    return (m.group(1) if m else "").strip()

def normalize_lines(section_text: str) -> list[str]:
    lines = [
        re.sub(r"\s+", " ", ln).strip()
        for ln in section_text.splitlines()
        if ln.strip()
    ]
    return lines

def parse_table(lines: list[str]) -> pd.DataFrame:
    rows = []
    i = 0
    while i < len(lines):
        ln = lines[i]
        if ln.lower().startswith("data valores") or "(70%)" in ln or "(30%)" in ln:
            i += 1
            continue

        if re.match(DATE_PATTERN, ln):
            date = ln.split()[0]
            rest = ln[len(date):].strip()
            parts = re.split(NUM_PATTERN, rest)
            nums = [p for p in parts if re.fullmatch(NUM_PATTERN, p)]
            desc = (parts[0] or "").strip()
            rows.append([
                date,
                desc,
                nums[0] if len(nums) > 0 else None,
                nums[1] if len(nums) > 1 else None,
                nums[2] if len(nums) > 2 else None,
                nums[3] if len(nums) > 3 else None,
            ])
            i += 1
            continue

        if ln.startswith("Aguardando"):
            desc = ln
            j = i + 1
            while j < len(lines) and not (
                re.match(DATE_PATTERN, lines[j]) or lines[j].startswith("Aguardando")
            ):
                desc += " " + lines[j]
                j += 1
            desc_clean = desc.replace("Aguardando ", "", 1)
            rows.append(["Aguardando", desc_clean, None, None, None, None])
            i = j
            continue

        rows.append([None, ln, None, None, None, None])
        i += 1

    df = pd.DataFrame(
        rows,
        columns=["Data","Valores (R$)","Principal (70%)","Contratual (30%)","Sucumbência","Total"],
    )
    return df.where(df.notna(), None)

def filter_most_recent(df: pd.DataFrame) -> pd.DataFrame:
    parsed = pd.to_datetime(df["Data"], format="%d/%m/%Y", errors="coerce")
    if parsed.notna().any():
        idx = parsed.idxmax()
        return df.loc[[idx]].reset_index(drop=True)
    return df.iloc[0:0].copy()

def br_money_to_float(x):
    """Converte '12.345,67' -> 12345.67"""
    if x is None:
        return None
    if isinstance(x, str):
        x = x.strip()
        if not x:
            return None
        x = x.replace(".", "").replace(",", ".")
    try:
        return float(x)
    except Exception:
        return None


def main():
    section_text = extract_section_text(PDF_PATH)
    if not section_text:
        raise RuntimeError("Seção '5. Comparativo dos Valores' não encontrada no PDF.")

    lines = normalize_lines(section_text)
    df = parse_table(lines)

    # --- Etapa 0: substituir None/NaN por "0,00" ---
    df = df.fillna("0,00")

    # --- Etapa 1: somar Contratual + Sucumbência ---
    df["_Contratual_num"]  = df["Contratual (30%)"].apply(br_money_to_float)
    df["_Sucumbencia_num"] = df["Sucumbência"].apply(br_money_to_float)
    df["Total"] = (
        df["_Contratual_num"].fillna(0) + df["_Sucumbencia_num"].fillna(0)
    )
    df["Total"] = df["Total"].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    df.drop(columns=["_Contratual_num", "_Sucumbencia_num"], inplace=True)

    # --- Etapa 2: manter apenas as 3 colunas desejadas ---
    df_final = df[["Contratual (30%)", "Sucumbência", "Total"]]

    # --- Etapa 3: pegar apenas a linha mais recente ---
    df_latest = filter_most_recent(df)
    df_latest = df_latest.fillna("0,00")

    df_latest["_Contratual_num"]  = df_latest["Contratual (30%)"].apply(br_money_to_float)
    df_latest["_Sucumbencia_num"] = df_latest["Sucumbência"].apply(br_money_to_float)
    df_latest["Total"] = (
        df_latest["_Contratual_num"].fillna(0) + df_latest["_Sucumbencia_num"].fillna(0)
    )
    df_latest["Total"] = df_latest["Total"].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    df_latest = df_latest[["Contratual (30%)", "Sucumbência", "Total"]]

    # --- Salvar ---
    Path(OUT_CSV).parent.mkdir(parents=True, exist_ok=True)
    df_final.to_csv(OUT_CSV, index=False, encoding="utf-8")
    df_final.to_excel(OUT_XLSX, index=False)

    df_latest.to_csv(OUT_CSV_LATEST, index=False, encoding="utf-8")
    df_latest.to_excel(OUT_XLSX_LATEST, index=False)

    print("Tabela completa (somente colunas desejadas):")
    print(df_final)
    print("\nLinha mais recente (somente colunas desejadas):")
    print(df_latest)
    print(f"\nArquivos salvos:\n- {OUT_CSV}\n- {OUT_XLSX}\n- {OUT_CSV_LATEST}\n- {OUT_XLSX_LATEST}")


if __name__ == "__main__":
    main()
