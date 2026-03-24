import os
import sys
import time
import threading
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

# Importações do Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# =========================================================
# REDIRECIONADOR DE LOGS (Prints para a Interface)
# =========================================================
class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        self.text_widget.insert(ctk.END, text)
        self.text_widget.see(ctk.END)  # Rola automaticamente para baixo

    def flush(self):
        pass

# =========================================================
# CLASSE PRINCIPAL DA INTERFACE
# =========================================================
class DigitadorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Robô Lançador de Notas - Versão Offline (CSV)")
        self.geometry("800x750")
        ctk.set_appearance_mode("System")
        
        # Variáveis de Controle
        self.csv_path = tk.StringVar(value="")
        self.driver = None

        # Configuração do Layout principal
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        self.criar_interface()
        
        # Redireciona os prints do console para a caixa de texto
        sys.stdout = TextRedirector(self.console_log)

    def criar_interface(self):
        # --- Seção do Arquivo CSV ---
        frame_csv = ctk.CTkFrame(self)
        frame_csv.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        
        ctk.CTkLabel(frame_csv, text="1. Arquivo de Notas (CSV)", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=5)
        
        ctk.CTkLabel(frame_csv, text="Selecione a planilha com as notas.\nO arquivo deve ter as colunas 'MATRICULA' e 'TOTAL'.", text_color="gray").grid(row=1, column=0, columnspan=2, pady=(0,10))

        btn_selecionar_csv = ctk.CTkButton(frame_csv, text="📂 Selecionar Arquivo CSV", command=self.selecionar_csv)
        btn_selecionar_csv.grid(row=2, column=0, columnspan=2, pady=10)
        
        self.lbl_csv = ctk.CTkLabel(frame_csv, text="Nenhum arquivo selecionado", text_color="gray")
        self.lbl_csv.grid(row=3, column=0, columnspan=2, pady=(0, 10))

        # --- Seção do Portal ---
        frame_portal = ctk.CTkFrame(self)
        frame_portal.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        ctk.CTkLabel(frame_portal, text="2. Configurações do Sistema de Destino", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=5)
        
        ctk.CTkLabel(frame_portal, text="URL do Sistema:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_url = ctk.CTkEntry(frame_portal, width=400, placeholder_text="Ex: https://atlasedu.com.br/notas...")
        self.entry_url.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # --- Seção de Estratégias ---
        frame_estrategia = ctk.CTkFrame(self)
        frame_estrategia.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        
        self.check_tabela = ctk.CTkCheckBox(frame_estrategia, text="Usar Estratégia de TABELA (Desmarque para busca individual)")
        self.check_tabela.select()
        self.check_tabela.pack(pady=10)

        # --- Botão de Iniciar ---
        self.btn_iniciar = ctk.CTkButton(self, text="🚀 INICIAR AUTOMAÇÃO", height=40, font=("Arial", 14, "bold"), command=self.iniciar_thread)
        self.btn_iniciar.grid(row=3, column=0, padx=20, pady=20)

        # --- Console de Logs ---
        ctk.CTkLabel(self, text="Console de Execução:").grid(row=4, column=0, padx=20, sticky="w")
        self.console_log = ctk.CTkTextbox(self, state="normal", wrap="word")
        self.console_log.grid(row=5, column=0, padx=20, pady=(0, 20), sticky="nsew")

    def selecionar_csv(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Arquivos CSV", "*.csv")])
        if arquivo:
            self.csv_path.set(arquivo)
            self.lbl_csv.configure(text=f"...{arquivo[-40:]}", text_color="green")

    def iniciar_thread(self):
        if not self.csv_path.get():
            messagebox.showwarning("Aviso", "Selecione o arquivo CSV com as notas!")
            return
        if not self.entry_url.get():
            messagebox.showwarning("Aviso", "Preencha a URL do sistema escolar!")
            return

        self.btn_iniciar.configure(state="disabled", text="Executando...")
        self.console_log.delete("1.0", ctk.END)
        
        thread = threading.Thread(target=self.rodar_automacao)
        thread.daemon = True
        thread.start()

    # =========================================================
    # LÓGICA DE AUTOMAÇÃO
    # =========================================================
    def ler_arquivo_csv(self, filepath):
        records = []
        try:
            # Detecta a codificação e o delimitador automaticamente (útil para arquivos do Excel BR)
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                conteudo = f.read(2048)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(conteudo)
                except csv.Error:
                    # Fallback padrão caso não consiga identificar
                    dialect = csv.excel
                
                reader = csv.DictReader(f, dialect=dialect)
                headers = [h.strip() for h in reader.fieldnames if h]
                
                print(f"📋 Cabeçalhos detectados no CSV: {headers}")
                
                if "MATRICULA" not in headers or "TOTAL" not in headers:
                    raise ValueError(f"O CSV precisa ter as colunas exatas 'MATRICULA' e 'TOTAL'. Colunas encontradas: {headers}")

                for row in reader:
                    # Limpa espaços em branco nas chaves e valores
                    clean_row = {k.strip(): str(v).strip() for k, v in row.items() if k}
                    records.append(clean_row)
            return records
        except UnicodeDecodeError:
            # Tenta ler com encoding ISO caso UTF-8 falhe (comum no Windows)
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
        
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver

    def rodar_automacao(self):
        try:
            print(f"📊 Lendo arquivo CSV: {self.csv_path.get()}...")
            records = self.ler_arquivo_csv(self.csv_path.get())
            
            if not records:
                raise RuntimeError("O arquivo CSV está vazio ou não pôde ser lido.")
            
            print(f"✅ Encontradas {len(records)} linhas com dados.\n")

            print("🌐 Abrindo navegador...")
            self.driver = self.make_driver()
            self.driver.get(self.entry_url.get().strip())

            # Pausa para o usuário fazer login
            messagebox.showinfo(
                "Ação Necessária", 
                "1. O navegador abriu.\n2. Faça login no portal escolar.\n3. Navegue até a tela onde as notas serão digitadas.\n4. Quando a tela estiver pronta, volte aqui e clique em OK."
            )

            print("\n🚀 Iniciando digitação automática...")
            
            use_table = self.check_tabela.get() == 1
            total_ok = 0
            not_found = []

            for i, row in enumerate(records, start=1):
                matricula = str(row.get("MATRICULA", "")).strip()
                nota_bruta = row.get("TOTAL", "")
                nota = self.normalizar_nota(nota_bruta)

                if nota == "" or matricula == "":
                    print(f"[{i:3d}] ⚠️ Matrícula ou nota vazia, pulando linha...")
                    continue

                print(f"[{i:3d}] 🔍 Processando: Matrícula {matricula} | Nota: {nota}")

                try:
                    if use_table:
                        ok = self.set_grade_in_table_row(matricula, nota)
                        if not ok:
                            not_found.append(matricula)
                        else:
                            total_ok += 1
                    else:
                        print("⚠️ Estratégia de busca unitária ativada. (Requer os seletores exatos configurados no código)")
                        
                except Exception as e:
                    print(f"[{i:3d}] 💥 ERRO para {matricula}: {e}")

            print("\n" + "="*60)
            print("📊 RESUMO FINAL:")
            print(f"   ✅ Notas digitadas: {total_ok}")
            if not_found:
                print(f"   ❌ Não encontradas na tela: {len(not_found)}")
                for m in not_found:
                    print(f"      - {m}")
            
            messagebox.showinfo("Finalizado", f"As notas foram preenchidas!\nSucesso: {total_ok}\nNão encontradas: {len(not_found)}\n\nVerifique no navegador e salve os dados.")

        except Exception as e:
            print(f"\n❌ ERRO CRÍTICO: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")
        finally:
            self.btn_iniciar.configure(state="normal", text="🚀 INICIAR AUTOMAÇÃO")

    def set_grade_in_table_row(self, matricula: str, nota: str) -> bool:
        # Mesma lógica do seu digitador original
        try:
            rows = self.driver.find_elements(By.CSS_SELECTOR, "tbody tr")
            for i, row in enumerate(rows):
                cells = row.find_elements(By.CSS_SELECTOR, "td")
                if len(cells) >= 4:
                    matricula_text = (cells[1].text or "").strip()
                    if matricula_text == str(matricula):
                        nota_input = cells[3].find_element(By.CSS_SELECTOR, "input.nota-input")
                        if nota_input:
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", nota_input)
                            time.sleep(0.3)
                            nota_input.clear()
                            
                            for char in nota:
                                nota_input.send_keys(char)
                                time.sleep(0.05)
                            
                            if nota_input.get_attribute('value') != nota:
                                self.driver.execute_script(f"arguments[0].value = '{nota}';", nota_input)
                                self.driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", nota_input)
                                
                            print(f"      ✅ OK {matricula} -> {nota}")
                            return True
            print(f"      ❌ Não encontrada na tabela: {matricula}")
            return False
        except Exception as e:
            print(f"      💥 Erro na tabela: {e}")
            return False

if __name__ == "__main__":
    app = DigitadorApp()
    app.mainloop()