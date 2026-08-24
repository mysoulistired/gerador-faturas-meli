"""
Interface gráfica para gerar as faturas a partir da Pré-fatura, usando
scripts/generate_faturas.py.

Roda com: python scripts/gui_faturas.py
"""
from __future__ import annotations

import datetime
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, scrolledtext, ttk

from generate_faturas import run_generation

BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) \
    else Path(__file__).resolve().parent.parent
SCRIPT_DIR = Path(__file__).resolve().parent
REQUIREMENTS = SCRIPT_DIR / "requirements.txt"

DEFAULT_PREFATURA = BASE_DIR / "Planilhas Revisar" / "Pré-fatura_Meli_Agosto - Q1.xlsx"
DEFAULT_TEMPLATE = BASE_DIR / "Planilhas Revisar" / "BETIM - BRMG02.xlsx"
DEFAULT_OUT_DIR = BASE_DIR / "Planilhas Revisar" / "Faturas Geradas"

BG = "#1e1e1e"
BG_PANEL = "#252526"
FG = "#d4d4d4"
FG_MUTED = "#9d9d9d"
ENTRY_BG = "#3c3c3c"
ACCENT = "#0e639c"
ACCENT_ACTIVE = "#1177bb"
BORDER = "#3c3c3c"


class FaturasGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Gerador de Faturas - Pré-fatura Mercado Livre")
        self.root.geometry("860x640")
        self.root.configure(bg=BG)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.running = False

        self._apply_dark_theme()
        self._build_form()
        self._build_buttons()
        self._build_console()

        self.root.after(100, self._poll_log_queue)

    def _apply_dark_theme(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure(".", background=BG, foreground=FG, fieldbackground=ENTRY_BG,
                         bordercolor=BORDER, lightcolor=BG, darkcolor=BG)
        style.configure("TFrame", background=BG)
        style.configure("TLabelframe", background=BG, foreground=FG, bordercolor=BORDER)
        style.configure("TLabelframe.Label", background=BG, foreground=FG)
        style.configure("TLabel", background=BG, foreground=FG)
        style.configure("TEntry", fieldbackground=ENTRY_BG, foreground=FG,
                         insertcolor=FG, bordercolor=BORDER)
        style.configure("TButton", background=ENTRY_BG, foreground=FG, bordercolor=BORDER,
                         focuscolor=BG, padding=6)
        style.map("TButton",
                  background=[("active", ACCENT_ACTIVE), ("disabled", BG_PANEL)],
                  foreground=[("disabled", FG_MUTED)])

    def _build_form(self):
        form = ttk.LabelFrame(self.root, text="Parâmetros")
        form.pack(fill="x", padx=10, pady=(10, 5))

        self.var_prefatura = tk.StringVar(value=str(DEFAULT_PREFATURA))
        self.var_template = tk.StringVar(value=str(DEFAULT_TEMPLATE))
        self.var_out_dir = tk.StringVar(value=str(DEFAULT_OUT_DIR))
        self.var_emissao = tk.StringVar(value=datetime.date.today().isoformat())
        self.var_periodo = tk.StringVar(value="")
        self.var_linehaul = tk.StringVar(value="")
        self.var_hub = tk.StringVar(value="")

        self._row_file(form, 0, "Pré-fatura (.xlsx):", self.var_prefatura, is_dir=False)
        self._row_file(form, 1, "Template (.xlsx):", self.var_template, is_dir=False)
        self._row_file(form, 2, "Pasta de saída:", self.var_out_dir, is_dir=True)

        row = 3
        ttk.Label(form, text="Emissão (AAAA-MM-DD):").grid(row=row, column=0, sticky="w", padx=5, pady=4)
        ttk.Entry(form, textvariable=self.var_emissao, width=20).grid(row=row, column=1, sticky="w", padx=5, pady=4)

        ttk.Label(form, text="Período (ex.: 202608Q1):").grid(row=row, column=2, sticky="w", padx=5, pady=4)
        ttk.Entry(form, textvariable=self.var_periodo, width=20).grid(row=row, column=3, sticky="w", padx=5, pady=4)

        row += 1
        ttk.Label(form, text="Line Haul N.:").grid(row=row, column=0, sticky="w", padx=5, pady=4)
        ttk.Entry(form, textvariable=self.var_linehaul, width=20).grid(row=row, column=1, sticky="w", padx=5, pady=4)

        ttk.Label(form, text="Hub específico (opcional):").grid(row=row, column=2, sticky="w", padx=5, pady=4)
        ttk.Entry(form, textvariable=self.var_hub, width=20).grid(row=row, column=3, sticky="w", padx=5, pady=4)

        for col in range(4):
            form.columnconfigure(col, weight=1)

    def _row_file(self, parent, row, label, var, is_dir):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=5, pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, columnspan=2, sticky="ew", padx=5, pady=4)
        cmd = (lambda: self._pick_dir(var)) if is_dir else (lambda: self._pick_file(var))
        ttk.Button(parent, text="Procurar...", command=cmd).grid(row=row, column=3, sticky="e", padx=5, pady=4)

    def _pick_file(self, var: tk.StringVar):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")], initialdir=str(BASE_DIR))
        if path:
            var.set(path)

    def _pick_dir(self, var: tk.StringVar):
        path = filedialog.askdirectory(initialdir=str(BASE_DIR))
        if path:
            var.set(path)

    def _build_buttons(self):
        bar = ttk.Frame(self.root)
        bar.pack(fill="x", padx=10, pady=5)

        self.btn_install = ttk.Button(bar, text="Instalar dependências", command=self.on_install)
        self.btn_install.pack(side="left", padx=(0, 8))
        if getattr(sys, "frozen", False):
            self.btn_install.state(["disabled"])

        self.btn_generate = ttk.Button(bar, text="Gerar Faturas", command=self.on_generate)
        self.btn_generate.pack(side="left")

        self.btn_clear = ttk.Button(bar, text="Limpar console", command=self.on_clear_console)
        self.btn_clear.pack(side="right")

    def _build_console(self):
        frame = ttk.LabelFrame(self.root, text="Console")
        frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        self.console = scrolledtext.ScrolledText(frame, state="disabled", bg="black", fg="#d0d0d0",
                                                   insertbackground="white", font=("Consolas", 10))
        self.console.pack(fill="both", expand=True)

    def on_install(self):
        if self.running:
            return
        cmd = [sys.executable, "-m", "pip", "install", "-r", str(REQUIREMENTS)]
        self._run_subprocess(cmd, "Instalando dependências")

    def on_generate(self):
        if self.running:
            return
        prefatura = self.var_prefatura.get().strip()
        template = self.var_template.get().strip()
        out_dir = self.var_out_dir.get().strip()
        emissao = self.var_emissao.get().strip()
        periodo = self.var_periodo.get().strip()
        linehaul = self.var_linehaul.get().strip()
        hub = self.var_hub.get().strip() or None

        faltando = [n for n, v in [("Pré-fatura", prefatura), ("Template", template),
                                    ("Pasta de saída", out_dir), ("Emissão", emissao),
                                    ("Período", periodo), ("Line Haul", linehaul)] if not v]
        if faltando:
            self._log(f"[!] Preencha os campos obrigatórios: {', '.join(faltando)}\n")
            return
        try:
            emissao_date = datetime.date.fromisoformat(emissao)
        except ValueError:
            self._log("[!] Data de emissão inválida. Use o formato AAAA-MM-DD.\n")
            return

        kwargs = dict(
            prefatura=Path(prefatura), template=Path(template), out_dir=Path(out_dir),
            emissao=emissao_date, periodo=periodo, linehaul=linehaul, hub=hub,
        )
        self._run_generation(kwargs)

    def on_clear_console(self):
        self.console.configure(state="normal")
        self.console.delete("1.0", tk.END)
        self.console.configure(state="disabled")

    def _start(self):
        self.running = True
        self.btn_install.state(["disabled"])
        self.btn_generate.state(["disabled"])

    def _run_subprocess(self, cmd: list[str], label: str):
        self._start()
        self._log(f"\n$ {' '.join(cmd)}\n")
        thread = threading.Thread(target=self._subprocess_worker, args=(cmd, label), daemon=True)
        thread.start()

    def _subprocess_worker(self, cmd: list[str], label: str):
        try:
            proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace",
                cwd=str(BASE_DIR),
            )
            for line in proc.stdout:
                self.log_queue.put(line)
            proc.wait()
            status = "concluído com sucesso" if proc.returncode == 0 else f"terminou com erro (código {proc.returncode})"
            self.log_queue.put(f"\n[{label}] {status}.\n")
        except Exception as exc:
            self.log_queue.put(f"\n[!] Falha ao executar: {exc}\n")
        finally:
            self.log_queue.put("__DONE__")

    def _run_generation(self, kwargs: dict):
        self._start()
        thread = threading.Thread(target=self._generation_worker, args=(kwargs,), daemon=True)
        thread.start()

    def _generation_worker(self, kwargs: dict):
        try:
            run_generation(**kwargs, log=lambda line: self.log_queue.put(line + "\n"))
        except Exception as exc:
            self.log_queue.put(f"\n[!] Falha ao gerar faturas: {exc}\n")
        finally:
            self.log_queue.put("__DONE__")

    def _poll_log_queue(self):
        try:
            while True:
                line = self.log_queue.get_nowait()
                if line == "__DONE__":
                    self.running = False
                    self.btn_install.state(["!disabled"])
                    self.btn_generate.state(["!disabled"])
                else:
                    self._log(line)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def _log(self, text: str):
        self.console.configure(state="normal")
        self.console.insert(tk.END, text)
        self.console.see(tk.END)
        self.console.configure(state="disabled")


def main():
    root = tk.Tk()
    FaturasGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
