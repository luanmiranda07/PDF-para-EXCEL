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
# Funções auxiliares
# ==========================

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def br_to_float(s: str) -> float:
    if not s:
        return 0.0
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

# ==========================
# Lógica de extração
# ==========================

def extrair_dados_pdf(caminho_pdf):
    dados = {
        "Arquivo": Path(caminho_pdf).name,
        "Nome": "",                  # <<< agora chama Nome
        "Número do processo": "",
        "Data encerramento": "",
        "CONTRATUAL": "",
        "Sucumbencia": "",
        "Total": ""
    }

    with pdfplumber.open(caminho_pdf) as pdf:
        texto = ""
        for page in pdf.pages:
            texto += (page.extract_text() or "") + "\n"

        # ---------- Nome (Parte Autora) ----------
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

        # ---------- Número do processo ----------
        mproc = re.search(r"(\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4})", texto)
        if mproc:
            dados["Número do processo"] = mproc.group(1)

        # ---------- Data de Encerramento ----------
        m = re.search(r"(\d{2}[/-]\d{2}[/-]\d{4}).{0,60}Encerramento\s+com\s+a\s+Libera",
                      texto, flags=re.IGNORECASE)
        if m:
            dados["Data encerramento"] = m.group(1)
        else:
            m = re.search(r"Encerramento\s+com\s+a\s+Libera[^\n\r]{0,60}?(\d{2}[/-]\d{2}[/-]\d{4})",
                          texto, flags=re.IGNORECASE)
            if m:
                dados["Data encerramento"] = m.group(1)

        # ---------- Previsto no Laudo ----------
        achou = False
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for table in tables:
                for row in table:
                    if row and any("Previstos no Laudo" in str(cell) for cell in row):
                        dados["CONTRATUAL"]  = row[2] if len(row) > 2 else ""
                        dados["Sucumbencia"] = row[3] if len(row) > 3 else ""

                        contratual = br_to_float(dados["CONTRATUAL"])
                        sucumb = br_to_float(dados["Sucumbencia"])
                        total = contratual + sucumb
                        dados["Total"] = (
                            f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        )
                        achou = True
                        break
                if achou:
                    break

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
        "Nome",                    # <<< aparece no Excel
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
