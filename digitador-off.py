import os
import sys
import time
import threading
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ============================================================
#  DESIGN SYSTEM — Paleta de Cores
# ============================================================
BG       = "#0d0d1a"
SURFACE  = "#13132a"
SURFACE2 = "#1c1c38"
BORDER   = "#2d2d52"
ACCENT   = "#7c6af7"
ACCENT_H = "#6c5ce7"
ACCENT_D = "#2e2870"
SUCCESS  = "#22c55e"
WARNING  = "#f59e0b"
ERROR    = "#ef4444"
TEXT     = "#e2e8f0"
TEXT2    = "#8892a4"
TEXT3    = "#3d4a5c"
CONSOLE  = "#08081a"

# ============================================================
#  CONSOLE COLORIDO
# ============================================================
class ColoredConsole:
    def __init__(self, textbox: ctk.CTkTextbox):
        self.tb = textbox
        inner = textbox._textbox
        inner.tag_configure("ok",      foreground="#4ade80")
        inner.tag_configure("err",     foreground="#f87171")
        inner.tag_configure("warn",    foreground="#fbbf24")
        inner.tag_configure("info",    foreground="#60a5fa")
        inner.tag_configure("dim",     foreground="#475569")
        inner.tag_configure("default", foreground="#c8d3e8")

    def write(self, text):
        inner = self.tb._textbox
        tag = "default"
        if any(x in text for x in ["✅", "OK →"]):
            tag = "ok"
        elif any(x in text for x in ["❌", "💥", "ERRO", "Error"]):
            tag = "err"
        elif any(x in text for x in ["⚠️", "⚠"]):
            tag = "warn"
        elif any(x in text for x in ["🚀", "📋", "🌐", "🔍", "📂", "📊"]):
            tag = "info"
        elif text.count("=") > 4 or text.count("-") > 4:
            tag = "dim"
        inner.insert("end", text, tag)
        inner.see("end")

    def flush(self):
        pass

# ============================================================
#  APLICATIVO PRINCIPAL
# ============================================================
class DigitadorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Digitador de Notas")
        self.geometry("860x800")
        self.minsize(720, 700)
        self.configure(fg_color=BG)

        self.csv_path  = tk.StringVar(value="")
        self.driver    = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self._build_header()
        self._build_config()
        self._build_action()
        self._build_console()
        self._build_statusbar()

        sys.stdout = ColoredConsole(self.console)

    # ─── HELPERS ─────────────────────────────────────────────

    def _step_badge(self, parent, num, row, col):
        ctk.CTkLabel(
            parent, text=num,
            font=("Inter", 10, "bold"),
            text_color="#a5b4fc",
            fg_color=ACCENT_D,
            width=24, height=24,
            corner_radius=12
        ).grid(row=row, column=col, padx=(16, 10), pady=(16, 0), sticky="n")

    def _card(self, parent, row):
        f = ctk.CTkFrame(parent, fg_color=SURFACE, border_color=BORDER,
                         border_width=1, corner_radius=12)
        f.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        f.grid_columnconfigure(1, weight=1)
        return f

    # ─── HEADER ──────────────────────────────────────────────

    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            hdr, text="⚡",
            font=("Inter", 28),
            text_color=ACCENT,
            fg_color=ACCENT_D,
            width=52, height=52,
            corner_radius=12
        ).grid(row=0, column=0, rowspan=2, padx=(20, 14), pady=16)

        ctk.CTkLabel(
            hdr, text="Digitador de Notas",
            font=("Inter", 19, "bold"),
            text_color=TEXT, anchor="w"
        ).grid(row=0, column=1, sticky="sw", pady=(16, 0))

        ctk.CTkLabel(
            hdr, text="Automação de lançamento via CSV  •  Versão 2.0",
            font=("Inter", 11),
            text_color=TEXT2, anchor="w"
        ).grid(row=1, column=1, sticky="nw", pady=(0, 16))

        ctk.CTkLabel(
            hdr, text="  OFFLINE  ",
            font=("Inter", 10, "bold"),
            text_color="#a5b4fc",
            fg_color=ACCENT_D,
            corner_radius=6
        ).grid(row=0, column=2, padx=(0, 20), pady=(18, 0), sticky="se")

    # ─── CONFIG ──────────────────────────────────────────────

    def _build_config(self):
        outer = ctk.CTkFrame(self, fg_color="transparent")
        outer.grid(row=1, column=0, sticky="ew", padx=20, pady=(16, 0))
        outer.grid_columnconfigure(0, weight=1)
        self._build_csv_card(outer)
        self._build_url_card(outer)
        self._build_strategy_card(outer)

    def _build_csv_card(self, parent):
        card = self._card(parent, 0)
        self._step_badge(card, "1", 0, 0)

        ctk.CTkLabel(card, text="Arquivo de Notas (CSV)",
                     font=("Inter", 13, "bold"), text_color=TEXT,
                     anchor="w").grid(row=0, column=1, sticky="w", pady=(14, 0))

        ctk.CTkLabel(card, text="Colunas obrigatórias: MATRICULA e TOTAL",
                     font=("Inter", 11), text_color=TEXT2,
                     anchor="w").grid(row=1, column=1, sticky="w", pady=(2, 10), columnspan=2)

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.grid(row=2, column=0, columnspan=3, sticky="ew", padx=16, pady=(0, 16))
        btn_row.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            btn_row, text="📂  Selecionar CSV",
            font=("Inter", 12, "bold"),
            fg_color=ACCENT, hover_color=ACCENT_H,
            text_color="white", corner_radius=8, height=36, width=170,
            command=self.selecionar_csv
        ).grid(row=0, column=0, padx=(0, 12))

        info = ctk.CTkFrame(btn_row, fg_color=SURFACE2, corner_radius=8)
        info.grid(row=0, column=1, sticky="ew")
        info.grid_columnconfigure(0, weight=1)

        self.lbl_csv = ctk.CTkLabel(
            info, text="Nenhum arquivo selecionado",
            font=("Inter", 11), text_color=TEXT3, anchor="w"
        )
        self.lbl_csv.grid(row=0, column=0, padx=12, pady=8, sticky="w")

    def _build_url_card(self, parent):
        card = self._card(parent, 1)
        self._step_badge(card, "2", 0, 0)

        ctk.CTkLabel(card, text="URL do Sistema Escolar",
                     font=("Inter", 13, "bold"), text_color=TEXT,
                     anchor="w").grid(row=0, column=1, sticky="w", pady=(14, 0))

        ctk.CTkLabel(card, text="Cole a URL da tela onde as notas serão digitadas",
                     font=("Inter", 11), text_color=TEXT2,
                     anchor="w").grid(row=1, column=1, sticky="w", pady=(2, 10), columnspan=2)

        self.entry_url = ctk.CTkEntry(
            card,
            placeholder_text="https://sistema.escolar.com.br/notas/lancamento...",
            font=("Inter", 12),
            fg_color=SURFACE2, border_color=BORDER,
            text_color=TEXT, placeholder_text_color=TEXT3,
            corner_radius=8, height=38
        )
        self.entry_url.grid(row=2, column=0, columnspan=3, sticky="ew",
                            padx=16, pady=(0, 16))

    def _build_strategy_card(self, parent):
        card = self._card(parent, 2)
        self._step_badge(card, "3", 0, 0)

        ctk.CTkLabel(card, text="Estratégia de Busca",
                     font=("Inter", 13, "bold"), text_color=TEXT,
                     anchor="w").grid(row=0, column=1, sticky="w", pady=(14, 0))

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.grid(row=1, column=0, columnspan=3, sticky="ew", padx=16, pady=(4, 16))

        self.switch_tabela = ctk.CTkSwitch(
            row,
            text="  Busca por TABELA  (recomendado — mais rápido)",
            font=("Inter", 12), text_color=TEXT,
            button_color=ACCENT, button_hover_color=ACCENT_H,
            progress_color=ACCENT_D,
            onvalue=1, offvalue=0
        )
        self.switch_tabela.select()
        self.switch_tabela.pack(anchor="w")

    # ─── BOTÃO + PROGRESS ────────────────────────────────────

    def _build_action(self):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=2, column=0, sticky="ew", padx=20, pady=16)
        frame.grid_columnconfigure(0, weight=1)

        self.btn_iniciar = ctk.CTkButton(
            frame,
            text="🚀   INICIAR AUTOMAÇÃO",
            font=("Inter", 14, "bold"),
            fg_color=ACCENT, hover_color=ACCENT_H,
            text_color="white",
            corner_radius=10, height=48,
            command=self.iniciar_thread
        )
        self.btn_iniciar.grid(row=0, column=0, sticky="ew")

        self.progress = ctk.CTkProgressBar(
            frame, fg_color=SURFACE2, progress_color=ACCENT,
            corner_radius=4, height=5
        )
        self.progress.set(0)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.progress.grid_remove()

    # ─── CONSOLE ─────────────────────────────────────────────

    def _build_console(self):
        outer = ctk.CTkFrame(self, fg_color=SURFACE, border_color=BORDER,
                             border_width=1, corner_radius=12)
        outer.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 0))
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        # Barra de título estilo macOS
        hdr = ctk.CTkFrame(outer, fg_color=SURFACE2, corner_radius=0, height=34)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)

        for color in ["#ef4444", "#f59e0b", "#22c55e"]:
            ctk.CTkLabel(hdr, text="●", text_color=color,
                         font=("Inter", 13)).pack(side="left", padx=(12, 2), pady=8)

        ctk.CTkLabel(hdr, text="Console de Execução",
                     font=("Inter", 11), text_color=TEXT2).pack(side="left", padx=8)

        ctk.CTkButton(
            hdr, text="Limpar",
            font=("Inter", 11),
            fg_color="transparent", hover_color=SURFACE,
            text_color=TEXT2, border_color=BORDER, border_width=1,
            corner_radius=6, height=24, width=56,
            command=lambda: self.console.delete("1.0", "end")
        ).pack(side="right", padx=10, pady=5)

        self.console = ctk.CTkTextbox(
            outer,
            fg_color=CONSOLE, text_color="#c8d3e8",
            font=("Courier", 11),
            corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=ACCENT_D,
            wrap="word", state="normal"
        )
        self.console.grid(row=1, column=0, sticky="nsew")

    # ─── STATUS BAR ──────────────────────────────────────────

    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0, height=32)
        bar.grid(row=4, column=0, sticky="ew", pady=(6, 0))
        bar.grid_propagate(False)

        self.lbl_status = ctk.CTkLabel(
            bar, text="● Pronto",
            font=("Inter", 11), text_color=TEXT2
        )
        self.lbl_status.pack(side="left", padx=16)

        self.lbl_stats = ctk.CTkLabel(
            bar, text="",
            font=("Inter", 11), text_color=TEXT2
        )
        self.lbl_stats.pack(side="right", padx=16)

    # ============================================================
    #  LÓGICA DE NEGÓCIO
    # ============================================================

    def selecionar_csv(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Arquivos CSV", "*.csv")])
        if arquivo:
            self.csv_path.set(arquivo)
            nome = os.path.basename(arquivo)
            try:
                records = self.ler_arquivo_csv(arquivo)
                count = len(records)
                self.lbl_csv.configure(
                    text=f"  📄  {nome}  •  {count} alunos",
                    text_color=SUCCESS
                )
                self.lbl_stats.configure(text=f"{count} registros carregados")
            except Exception:
                self.lbl_csv.configure(text=f"  📄  {nome}", text_color=TEXT)

    def iniciar_thread(self):
        if not self.csv_path.get():
            messagebox.showwarning("Aviso", "Selecione o arquivo CSV com as notas!")
            return
        if not self.entry_url.get():
            messagebox.showwarning("Aviso", "Preencha a URL do sistema escolar!")
            return

        self.btn_iniciar.configure(state="disabled", text="⏳   Executando...",
                                   fg_color=ACCENT_D)
        self.console.delete("1.0", "end")
        self.lbl_status.configure(text="● Executando", text_color=WARNING)
        self.progress.grid()
        self.progress.configure(mode="indeterminate")
        self.progress.start()

        thread = threading.Thread(target=self.rodar_automacao)
        thread.daemon = True
        thread.start()

    def _reset_ui(self, success=True):
        self.btn_iniciar.configure(state="normal", text="🚀   INICIAR AUTOMAÇÃO",
                                   fg_color=ACCENT)
        self.progress.stop()
        self.progress.grid_remove()
        self.lbl_status.configure(
            text="● Concluído" if success else "● Erro",
            text_color=SUCCESS if success else ERROR
        )

    def ler_arquivo_csv(self, filepath):
        records = []
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                conteudo = f.read(2048)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(conteudo)
                except csv.Error:
                    dialect = csv.excel
                reader = csv.DictReader(f, dialect=dialect)
                headers = [h.strip() for h in reader.fieldnames if h]
                print(f"📋 Cabeçalhos detectados: {headers}")
                if "MATRICULA" not in headers or "TOTAL" not in headers:
                    raise ValueError(f"Colunas necessárias: MATRICULA e TOTAL. Encontradas: {headers}")
                for row in reader:
                    clean_row = {k.strip(): str(v).strip() for k, v in row.items() if k}
                    records.append(clean_row)
            return records
        except UnicodeDecodeError:
            with open(filepath, 'r', encoding='iso-8859-1') as f:
                reader = csv.DictReader(f, delimiter=';')
                for row in reader:
                    clean_row = {k.strip(): str(v).strip() for k, v in row.items() if k}
                    records.append(clean_row)
            return records

    def normalizar_nota(self, valor):
        if valor is None:
            return ""
        valor_str = str(valor)
        if "." in valor_str and valor_str.replace(".", "").replace("-", "").isdigit():
            return valor_str.replace(".", ",")
        return valor_str

    def make_driver(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=options
        )
        driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        return driver

    def rodar_automacao(self):
        try:
            print(f"📊 Lendo arquivo: {self.csv_path.get()}...")
            records = self.ler_arquivo_csv(self.csv_path.get())
            if not records:
                raise RuntimeError("O arquivo CSV está vazio ou não pôde ser lido.")
            print(f"✅ {len(records)} registros encontrados.\n")

            print("🌐 Abrindo navegador...")
            self.driver = self.make_driver()
            self.driver.get(self.entry_url.get().strip())

            messagebox.showinfo(
                "Ação Necessária",
                "1. O navegador abriu.\n"
                "2. Faça login no portal escolar.\n"
                "3. Navegue até a tela de lançamento de notas.\n"
                "4. Quando estiver pronto, clique em OK."
            )

            print("🚀 Iniciando digitação automática...\n")
            use_table = self.switch_tabela.get() == 1
            total_ok, not_found = 0, []

            for i, row in enumerate(records, start=1):
                matricula = str(row.get("MATRICULA", "")).strip()
                nota = self.normalizar_nota(row.get("TOTAL", ""))
                if not nota or not matricula:
                    print(f"[{i:3d}] ⚠️  Linha vazia, pulando...")
                    continue
                print(f"[{i:3d}] 🔍 Matrícula {matricula}  →  Nota {nota}")
                try:
                    if use_table:
                        ok = self.set_grade_in_table_row(matricula, nota)
                        if not ok:
                            not_found.append(matricula)
                        else:
                            total_ok += 1
                    else:
                        print("⚠️  Estratégia individual: configure os seletores no código.")
                except Exception as e:
                    print(f"[{i:3d}] 💥 ERRO para {matricula}: {e}")

            print("\n" + "=" * 55)
            print("📊 RESUMO FINAL")
            print(f"   ✅ Digitadas com sucesso : {total_ok}")
            if not_found:
                print(f"   ❌ Não encontradas      : {len(not_found)}")
                for m in not_found:
                    print(f"      — {m}")
            print("=" * 55)

            self.after(0, lambda: self.lbl_stats.configure(
                text=f"✅ Sucesso: {total_ok}   ❌ Não enc.: {len(not_found)}"
            ))

            messagebox.showinfo(
                "Finalizado ✅",
                f"Lançamento concluído!\n\n"
                f"✅ Sucesso: {total_ok}\n"
                f"❌ Não encontradas: {len(not_found)}\n\n"
                "Verifique no navegador e salve os dados."
            )
            self.after(0, lambda: self._reset_ui(success=True))

        except Exception as e:
            print(f"\n❌ ERRO CRÍTICO: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")
            self.after(0, lambda: self._reset_ui(success=False))

    def set_grade_in_table_row(self, matricula: str, nota: str) -> bool:
        try:
            rows = self.driver.find_elements(By.CSS_SELECTOR, "tbody tr")
            for row in rows:
                cells = row.find_elements(By.CSS_SELECTOR, "td")
                if len(cells) >= 4:
                    matricula_text = (cells[1].text or "").strip()
                    if matricula_text == str(matricula):
                        nota_input = cells[3].find_element(By.CSS_SELECTOR, "input.nota-input")
                        if nota_input:
                            self.driver.execute_script(
                                "arguments[0].scrollIntoView({block: 'center'});", nota_input
                            )
                            time.sleep(0.3)
                            nota_input.clear()
                            for char in nota:
                                nota_input.send_keys(char)
                                time.sleep(0.05)
                            if nota_input.get_attribute('value') != nota:
                                self.driver.execute_script(
                                    f"arguments[0].value = '{nota}';", nota_input
                                )
                                self.driver.execute_script(
                                    "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));",
                                    nota_input
                                )
                            print(f"      ✅ OK → {matricula}: {nota}")
                            return True
            print(f"      ❌ Não encontrada na tabela: {matricula}")
            return False
        except Exception as e:
            print(f"      💥 Erro na tabela: {e}")
            return False


if __name__ == "__main__":
    app = DigitadorApp()
    app.mainloop()