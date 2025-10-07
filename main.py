import os
import glob
import re
import unicodedata
from pathlib import Path

import pdfplumber
import pandas as pd

# ==== TKINTER UI ====
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ==========================
# Funções auxiliares (GERAIS)
# ==========================

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def br_to_float(s: str) -> float:
    """Converte '12.345,67' -> 12345.67. Vazio/erro -> 0.0"""
    if not s:
        return 0.0
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

def float_to_br(x: float) -> str:
    """Converte 12345.67 -> '12.345,67' (PT-BR)."""
    try:
        s = f"{float(x):,.2f}"          # '12,345.67'
        s = s.replace(",", "X")         # '12X345.67'
        s = s.replace(".", ",")         # '12X345,67'
        s = s.replace("X", ".")         # '12.345,67'
        return s
    except:
        return ""

# ==========================
# Helpers específicos da Seção 5 - COMPARATIVO DOS VALORES
# ==========================

HEADER_PATTERN = r"5\.\s*Comparativo dos Valores"
NEXT_HEADER    = r"\n\s*6\.\s*Observações|\Z"
NUM_PATTERN    = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
DATE_PATTERN   = r"\d{2}/\d{2}/\d{4}"

def extract_section5_from_text(full_text: str) -> str:
    """Pega o texto entre '5. Comparativo dos Valores' e '6. Observações'."""
    m = re.search(HEADER_PATTERN + r"(.*?)" + NEXT_HEADER, full_text, flags=re.S)
    return (m.group(1) if m else "").strip()

def normalize_lines(section_text: str) -> list[str]:
    """Quebra em linhas e compacta espaços."""
    lines = [
        re.sub(r"\s+", " ", ln).strip()
        for ln in section_text.splitlines()
        if ln.strip()
    ]
    return lines

def parse_table(lines: list[str]) -> pd.DataFrame:
    """
    Constrói a tabela com colunas:
      Data | Valores (R$) | Principal (70%) | Contratual (30%) | Sucumbência | Total
    """
    rows = []
    i = 0
    while i < len(lines):
        ln = lines[i]

        # pula cabeçalhos
        if ln.lower().startswith("data valores") or "(70%)" in ln or "(30%)" in ln:
            i += 1
            continue

        if re.match(DATE_PATTERN, ln):  # caso com data dd/mm/yyyy
            date = ln.split()[0]
            rest = ln[len(date):].strip()

            # captura todos os números na linha
            nums = re.findall(NUM_PATTERN, rest)

            # descrição = texto antes do 1º número (se houver)
            first_num_match = re.search(NUM_PATTERN, rest)
            desc = rest[:first_num_match.start()].strip() if first_num_match else rest.strip()

            # Preenche da ESQUERDA para DIREITA
            principal = contratual = sucumbencia = total = "0,00"

            if len(nums) >= 1:
                principal = nums[0]  # Primeiro número -> Principal (70%)
            if len(nums) >= 2:
                contratual = nums[1]  # Segundo número -> Contratual (30%)
            if len(nums) >= 3:
                sucumbencia = nums[2]  # Terceiro número -> Sucumbência
            if len(nums) >= 4:
                total = nums[3]  # Quarto número -> Total

            rows.append([date, desc, principal, contratual, sucumbencia, total])
            i += 1
            continue

        if ln.startswith("Aguardando"):  # bloco "Aguardando ..."
            desc = ln
            j = i + 1
            while j < len(lines) and not (
                re.match(DATE_PATTERN, lines[j]) or lines[j].startswith("Aguardando")
            ):
                desc += " " + lines[j]
                j += 1
            desc_clean = desc.replace("Aguardando ", "", 1)
            rows.append(["Aguardando", desc_clean, "0,00", "0,00", "0,00", "0,00"])
            i = j
            continue

        # fallback
        rows.append([None, ln, "0,00", "0,00", "0,00", "0,00"])
        i += 1

    df = pd.DataFrame(
        rows,
        columns=["Data", "Valores (R$)", "Principal (70%)", "Contratual (30%)", "Sucumbência", "Total"]
    )
    return df

def filter_most_recent(df: pd.DataFrame) -> pd.DataFrame:
    """Retorna um DF com apenas a linha da data mais recente (ignora 'Aguardando')."""
    # Filtra apenas linhas com datas válidas
    valid_dates = df[df["Data"].str.match(DATE_PATTERN, na=False)].copy()
    
    if valid_dates.empty:
        return df.iloc[0:0].copy()
    
    # Converte para datetime e pega a mais recente
    valid_dates["Data_dt"] = pd.to_datetime(valid_dates["Data"], format="%d/%m/%Y", errors="coerce")
    idx = valid_dates["Data_dt"].idxmax()
    
    return df.loc[[idx]].reset_index(drop=True)

def processar_secao5_comparativo(texto: str):
    """Processa a seção 5 do PDF e retorna os dados do comparativo"""
    section5_text = extract_section5_from_text(texto)
    
    if not section5_text:
        return "0,00", "0,00", "0,00"
    
    lines = normalize_lines(section5_text)
    df_sec5 = parse_table(lines)
    df_latest = filter_most_recent(df_sec5)

    if not df_latest.empty:
        contratual_txt = df_latest.at[0, "Contratual (30%)"] or "0,00"
        sucumb_txt = df_latest.at[0, "Sucumbência"] or "0,00"

        contratual_val = br_to_float(contratual_txt)
        sucumb_val = br_to_float(sucumb_txt)
        total_val = contratual_val + sucumb_val

        return contratual_txt, sucumb_txt, float_to_br(total_val)
    
    return "0,00", "0,00", "0,00"

# ==========================
# Lógica de extração PRINCIPAL
# ==========================

def extrair_dados_pdf(caminho_pdf):
    dados = {
        "Arquivo": Path(caminho_pdf).name,
        "Nome": "",
        "Número do processo": "",
        "Data encerramento": "",
        "CONTRATUAL": "",
        "Sucumbencia": "",
        "Total": ""
    }

    with pdfplumber.open(caminho_pdf) as pdf:
        # — 1) extrai texto de todas as páginas —
        texto = ""
        for page in pdf.pages:
            texto += (page.extract_text() or "") + "\n"

        # — 2) Nome (Parte Autora) —
        m_aut = re.search(
            r"parte\s+autor(?:a|o)(?:\s*\(o\))?\s*[:\-–]?\s*([^\n\r|]+)",
            texto, flags=re.IGNORECASE
        )
        if m_aut:
            dados["Nome"] = m_aut.group(1).strip()

        if not dados["Nome"]:
            for page in pdf.pages:
                tables = page.extract_tables() or []
                achou_pa = False
                for table in tables:
                    for row in table:
                        if not row:
                            continue
                        for i, cell in enumerate(row):
                            txt = str(cell) if cell else ""
                            if "parte autora" in _norm(txt):
                                nome = ""
                                if ":" in txt:
                                    nome = txt.split(":", 1)[1].strip()
                                if not nome and i + 1 < len(row):
                                    nome = (str(row[i+1]) if row[i+1] else "").strip()
                                if nome:
                                    dados["Nome"] = nome
                                    achou_pa = True
                                    break
                        if achou_pa:
                            break
                    if achou_pa:
                        break

        # — 3) Número do processo —
        mproc = re.search(r"(\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4})", texto)
        if mproc:
            dados["Número do processo"] = mproc.group(1)

        # — 4) Data de Encerramento —
        m = re.search(r"(\d{2}[/-]\d{2}[/-]\d{4}).{0,60}Encerramento\s+com\s+a\s+Libera",
                      texto, flags=re.IGNORECASE)
        if m:
            dados["Data encerramento"] = m.group(1)
        else:
            m = re.search(r"Encerramento\s+com\s+a\s+Libera[^\n\r]{0,60}?(\d{2}[/-]\d{2}[/-]\d{4})",
                          texto, flags=re.IGNORECASE)
            if m:
                dados["Data encerramento"] = m.group(1)

        # — 5) Seção 5: Comparativo dos Valores —
        contratual, sucumbencia, total = processar_secao5_comparativo(texto)
        dados["CONTRATUAL"] = contratual
        dados["Sucumbencia"] = sucumbencia
        dados["Total"] = total

    return dados

# ==========================
# Funções usadas pela UI
# ==========================

def extrair_em_lote(lista_caminhos):
    resultados = []
    for p in lista_caminhos:
        try:
            d = extrair_dados_pdf(p)
            registro = {
                "Arquivo": d.get("Arquivo", Path(p).name),
                "Nome": d.get("Nome", ""),
                "Número do processo": d.get("Número do processo", ""),
                "Data Encerramento": d.get("Data encerramento", ""),
                "Total": d.get("Total", ""),
                "CONTRATUAL": d.get("CONTRATUAL", ""),
                "Sucumbencia": d.get("Sucumbencia", ""),
                "Erro": ""
            }
            resultados.append(registro)
        except Exception as e:
            resultados.append({
                "Arquivo": Path(p).name,
                "Nome": "",
                "Número do processo": "",
                "Data Encerramento": "",
                "Total": "",
                "CONTRATUAL": "",
                "Sucumbencia": "",
                "Erro": str(e),
            })
    return resultados

def gerar_excel_em_arquivo_lote(result_rows, caminho_saida):
    df = pd.DataFrame(result_rows)
    colunas_excel = [
        "Arquivo",
        "Nome",
        "Número do processo",
        "Data Encerramento",
        "CONTRATUAL",
        "Sucumbencia",
        "Total",
    ]
    for c in colunas_excel:
        if c not in df.columns:
            df[c] = ""
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        df[colunas_excel].to_excel(writer, index=False, sheet_name="Dados")

# ==========================
# Interface Tkinter
# ==========================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Extrair PDF para EXCEL (Lote)")
        self.geometry("1000x480")
        self.minsize(950, 440)

        self.current_files = []
        self.result_rows = []
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=12)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Arquivos PDF:").pack(side=tk.LEFT)
        ttk.Button(top, text="Abrir… (múltiplos)", command=self.on_open_many).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(top, text="Abrir pasta…", command=self.on_open_folder).pack(side=tk.LEFT, padx=(8, 0))

        self.recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Incluir subpastas", variable=self.recursive_var).pack(side=tk.LEFT, padx=12)

        mid = ttk.Frame(self, padding=(6, 3))
        mid.pack(fill=tk.BOTH, expand=True)

        cols = ("Arquivo", "Nome", "Número do processo", "Data Encerramento", "Total")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=12)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=180 if c not in ("Arquivo", "Nome") else 220, anchor=tk.W)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        actions = ttk.Frame(self, padding=12)
        actions.pack(fill=tk.X)
        ttk.Button(actions, text="Processar", command=self.on_process).pack(side=tk.LEFT)
        ttk.Button(actions, text="Salvar Excel", command=self.on_save_excel).pack(side=tk.LEFT, padx=8)
        ttk.Button(actions, text="Limpar", command=self.on_clear).pack(side=tk.LEFT)
        ttk.Button(actions, text="Remover seleção", command=self.on_remove_selected).pack(side=tk.LEFT, padx=8)

        self.status_var = tk.StringVar(value="Pronto.")
        ttk.Label(self, textvariable=self.status_var, anchor=tk.W).pack(fill=tk.X, padx=12, pady=(0,12))

    # ===== Ações =====
    def on_open_many(self):
        paths = filedialog.askopenfilenames(title="Selecionar arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if paths:
            for p in paths:
                if p not in self.current_files:
                    self.current_files.append(p)
            self.status_var.set(f"{len(paths)} arquivo(s) adicionado(s). Total: {len(self.current_files)}.")

    def on_open_folder(self):
        folder = filedialog.askdirectory(title="Selecionar pasta")
        if not folder:
            return
        pattern = "**/*.pdf" if self.recursive_var.get() else "*.pdf"
        encontrados = glob.glob(os.path.join(folder, pattern), recursive=self.recursive_var.get())
        if not encontrados:
            messagebox.showinfo("Info", "Nenhum arquivo .pdf encontrado.")
            return
        for p in encontrados:
            if p not in self.current_files:
                self.current_files.append(p)
        self.status_var.set(f"{len(encontrados)} arquivo(s) da pasta adicionados. Total: {len(self.current_files)}.")

    def on_process(self):
        if not self.current_files:
            messagebox.showwarning("Aviso", "Selecione arquivos ou uma pasta primeiro.")
            return
        self.status_var.set("Processando… aguarde.")
        self.update_idletasks()

        self.result_rows = extrair_em_lote(self.current_files)

        for i in self.tree.get_children():
            self.tree.delete(i)
        for row in self.result_rows:
            self.tree.insert("", tk.END,
                values=(
                    row.get("Arquivo",""),
                    row.get("Nome",""),
                    row.get("Número do processo",""),
                    row.get("Data Encerramento",""),
                    row.get("Total",""),
                )
            )

        ok = sum(1 for r in self.result_rows if any([r.get("Número do processo"), r.get("Data Encerramento"), r.get("Total")]))
        falhas = sum(1 for r in self.result_rows if r.get("Erro"))
        self.status_var.set(f"Concluído: {ok} com dados, {falhas} com erro, total {len(self.result_rows)}.")

    def on_save_excel(self):
        if not self.result_rows:
            messagebox.showwarning("Aviso", "Nada para salvar: processe arquivos antes.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Planilha do Excel", "*.xlsx")],
                                            title="Salvar como",
                                            initialfile="dados_lote.xlsx")
        if path:
            try:
                gerar_excel_em_arquivo_lote(self.result_rows, path)
                self.status_var.set(f"Excel salvo em: {path}")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível salvar o Excel.\n\n{e}")

    def on_remove_selected(self):
        selecionados = self.tree.selection()
        if not selecionados:
            return
        nomes = {self.tree.item(i, "values")[0] for i in selecionados}
        self.current_files = [p for p in self.current_files if os.path.basename(p) not in nomes]
        for i in selecionados:
            self.tree.delete(i)
        if self.result_rows:
            self.result_rows = [r for r in self.result_rows if r.get("Arquivo") not in nomes]
        self.status_var.set(f"Removidos {len(nomes)} arquivo(s). Restam {len(self.current_files)}.")

    def on_clear(self):
        self.current_files = []
        self.result_rows = []
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.status_var.set("Pronto.")

if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = App()
    app.mainloop()