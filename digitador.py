import os
import time
from typing import Dict, Any, List

import gspread
from google.oauth2.service_account import Credentials

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# =========================================================
# CONFIGURAÇÕES GERAIS
# =========================================================

# (1) Google Sheets
SHEET_ID = "1Lb1Wf5AW8VbfRrzOMHDELMc3gvC6RLogQliYKCcMyNY"   # ex.: "1AbCDeF..."
WORKSHEET_NAME = "Notas"                      # nome da aba na planilha
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")

# Cabeçalhos exatamente como aparecem na planilha
HEADER_MATRICULA = "Número de identificação"
HEADER_NOTA = "TOTAL"

# (2) Sistema de destino (Atlas)
TARGET_URL = "https://atlasedu.com.br/notas/lancamento/lancar-nota"  # coloque a URL correta
LOGIN_NECESSARIO = False
LOGIN_USUARIO = os.environ.get("PORTAL_USER", "seu_usuario")
LOGIN_SENHA = os.environ.get("PORTAL_PASS", "sua_senha")

# Configurações do navegador
USAR_PERFIL_PERSISTENTE = True  # True = mantém login/cookies; False = navegador limpo a cada execução

# Se a área de digitação estiver DENTRO de um iframe, informe o seletor dele
USE_IFRAME = False
IFRAME_CSS = "iframe#iframe-notas"  # ajuste se necessário

# (3) Seletores (ADAPTE aos da sua página)
# Você pode usar:
#   - Estratégia A (BUSCA por matrícula + campo de nota),
#   - Estratégia B (TABELA de linhas: matrícula e input de nota na mesma linha).
USE_TABLE_STRATEGY = True  # True = usa tabela (Estratégia B); False = usa busca (Estratégia A)

SELECTORS: Dict[str, Any] = {
    # -------- Estratégia A: BUSCA --------
    "search_input": {"by": By.CSS_SELECTOR, "value": "input[name='matricula']"},
    "search_button": {"by": By.CSS_SELECTOR, "value": "button#buscar"},
    "grade_input": {"by": By.CSS_SELECTOR, "value": "input[name='nota']"},
    "save_button": {"by": By.CSS_SELECTOR, "value": "button#salvar-nota"},

    # -------- Estratégia B: TABELA --------
    "rows_container": {"by": By.CSS_SELECTOR, "value": "table#alunos tbody"},
    "row_items": {"by": By.CSS_SELECTOR, "value": "table#alunos tbody tr"},
    "row_matricula_cell": {"by": By.CSS_SELECTOR, "value": "td.mat"},
    "row_grade_input": {"by": By.CSS_SELECTOR, "value": "input.nota"},
    "row_save_button": {"by": By.CSS_SELECTOR, "value": "button.salvar"},  # se não houver, será ignorado
}

# Tempo padrão de espera explícita
TIMEOUT = 20

# Pequena pausa entre ações para evitar bloqueios (ajuste se necessário)
ACTION_SLEEP = 0.3


# =========================================================
# FUNÇÕES: Google Sheets
# =========================================================
def open_worksheet(sheet_id: str, worksheet_name: str):
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    sh = client.open_by_key(sheet_id)
    return sh.worksheet(worksheet_name)

def read_rows_as_dicts(ws) -> List[Dict[str, Any]]:
    records = ws.get_all_records()
    if not records:
        raise RuntimeError("Planilha vazia ou não acessível.")
    
    headers = set(records[0].keys())
    needed = {HEADER_MATRICULA, HEADER_NOTA}
    missing = needed - headers
    if missing:
        raise RuntimeError(f"Colunas ausentes na planilha: {missing}. Cabeçalhos encontrados: {list(headers)}")
    
    # Debug: mostra as primeiras linhas para entender o formato dos dados
    print(f"📊 DEBUG Planilha:")
    print(f"   - Cabeçalhos: {list(headers)}")
    print(f"   - Total de linhas: {len(records)}")
    print(f"   - Primeiras 3 linhas:")
    
    for i, record in enumerate(records[:3]):
        print(f"   Linha {i+1}:")
        print(f"     - {HEADER_MATRICULA}: '{record[HEADER_MATRICULA]}' (tipo: {type(record[HEADER_MATRICULA])})")
        print(f"     - {HEADER_NOTA}: '{record[HEADER_NOTA]}' (tipo: {type(record[HEADER_NOTA])})")
    
    return records

def normalizar_nota(valor):
    """FORÇA o formato brasileiro com vírgula para decimais."""
    if valor is None:
        return ""
    
    # Converte para string
    valor_str = str(valor)
    print(f"      🔥 NORMALIZAÇÃO BRUTA: '{valor}' -> '{valor_str}'")
    
    # Se for número decimal (float), FORÇA vírgula
    if isinstance(valor, float):
        if valor == int(valor):  # Número inteiro
            return str(int(valor))
        else:  # Número decimal - FORÇA vírgula
            # Converte para string e substitui ponto por vírgula
            resultado = str(valor).replace(".", ",")
            print(f"      ✅ FORÇADO: {valor} -> '{resultado}'")
            return resultado
    
    # Se for string, verifica se tem ponto decimal e converte para vírgula
    if "." in valor_str and valor_str.replace(".", "").replace("-", "").isdigit():
        resultado = valor_str.replace(".", ",")
        print(f"      ✅ STRING CONVERTIDA: '{valor_str}' -> '{resultado}'")
        return resultado
    
    # Se já tem vírgula, mantém como está
    if "," in valor_str:
        print(f"      ✅ JÁ TEM VÍRGULA: '{valor_str}'")
        return valor_str
    
    # Para outros casos, retorna como está
    return valor_str


# =========================================================
# FUNÇÕES: Selenium/WebDriver
# =========================================================
def make_driver(headless: bool = False):
    options = webdriver.ChromeOptions()
    
    # Configurações para manter dados do usuário
    if USAR_PERFIL_PERSISTENTE:
        import os
        user_data_dir = os.path.expanduser("~/chrome-automation-profile")
        
        # Cria o diretório se não existir
        if not os.path.exists(user_data_dir):
            os.makedirs(user_data_dir)
        
        # Usa um perfil persistente do Chrome
        options.add_argument(f"--user-data-dir={user_data_dir}")
        options.add_argument("--profile-directory=Default")
        print("🔐 Usando perfil persistente - login e cookies serão mantidos")
    else:
        print("🧹 Usando navegador limpo - dados serão perdidos a cada execução")
    
    # Outras configurações
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1366,768")
    
    # Evita mensagens de "Chrome está sendo controlado por software automatizado"
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    
    # Remove a mensagem de automação
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

def wait_el(driver, by, value, timeout=TIMEOUT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def wait_clickable(driver, by, value, timeout=TIMEOUT):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))

def do_login_if_needed(driver):
    if not LOGIN_NECESSARIO:
        return
    # Exemplo genérico de login — ajuste os seletores da sua tela
    user_input = wait_el(driver, By.CSS_SELECTOR, "input[name='usuario']")
    pass_input = wait_el(driver, By.CSS_SELECTOR, "input[name='senha']")
    login_btn  = wait_clickable(driver, By.CSS_SELECTOR, "button[type='submit']")
    user_input.clear(); user_input.send_keys(LOGIN_USUARIO)
    pass_input.clear(); pass_input.send_keys(LOGIN_SENHA)
    login_btn.click()
    # Aguarde um indicador da tela pós-login (ajuste para algo da sua aplicação)
    wait_el(driver, By.CSS_SELECTOR, "nav .menu-principal")

def switch_into_iframe_if_needed(driver):
    if USE_IFRAME:
        iframe = wait_el(driver, By.CSS_SELECTOR, IFRAME_CSS)
        driver.switch_to.frame(iframe)

def switch_out_iframe(driver):
    if USE_IFRAME:
        driver.switch_to.default_content()


# =========================================================
# ESTRATÉGIAS DE LANÇAMENTO
# =========================================================
def set_grade_using_search(driver, matricula: str, nota: str):
    search_input = wait_el(driver, **SELECTORS["search_input"])
    btn = wait_clickable(driver, **SELECTORS["search_button"])
    search_input.clear(); search_input.send_keys(str(matricula))
    time.sleep(ACTION_SLEEP)
    btn.click()
    time.sleep(ACTION_SLEEP)

    grade_input = wait_el(driver, **SELECTORS["grade_input"])
    grade_input.clear(); grade_input.send_keys(nota)
    time.sleep(ACTION_SLEEP)

    try:
        save_btn = wait_clickable(driver, **SELECTORS["save_button"])
        save_btn.click()
        time.sleep(ACTION_SLEEP)
    except Exception:
        # Caso não exista botão salvar (auto-save)
        pass

def set_grade_in_table_row(driver, matricula: str, nota: str) -> bool:
    """
    Procura a linha com a matrícula e insere a nota no input correspondente.
    Retorna True se encontrou/salvou; False se não encontrou.
    """
    try:
        print(f"      🔍 Procurando matrícula: {matricula}")
        
        # Procura por todas as linhas da tabela
        rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
        print(f"      📊 Encontradas {len(rows)} linhas na tabela")
        
        for i, row in enumerate(rows):
            try:
                # Cada linha tem múltiplas colunas (td)
                cells = row.find_elements(By.CSS_SELECTOR, "td")
                
                if len(cells) >= 4:  # Precisamos de pelo menos 4 colunas
                    # Coluna 1 (índice 0): Nome do aluno
                    # Coluna 2 (índice 1): Matrícula
                    # Coluna 3 (índice 2): Chamada
                    # Coluna 4 (índice 3): Nota
                    
                    matricula_cell = cells[1]  # 2ª coluna
                    matricula_text = (matricula_cell.text or "").strip()
                    
                    print(f"      📝 Linha {i}: Matrícula encontrada: '{matricula_text}'")
                    
                    # Verifica se esta linha contém a matrícula que procuramos
                    if matricula_text == str(matricula):
                        print(f"      🎯 Matrícula {matricula} encontrada na linha {i}")
                        
                        # Coluna da nota (4ª coluna, índice 3)
                        nota_cell = cells[3]
                        
                        # Procura o input de nota dentro desta célula
                        nota_input = nota_cell.find_element(By.CSS_SELECTOR, "input.nota-input")
                        
                        if nota_input:
                            print(f"      🎯 Input de nota encontrado na linha {i}")
                            
                            # Rola para o input
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", nota_input)
                            time.sleep(ACTION_SLEEP)
                            
                            # Debug antes de inserir
                            print(f"      🔍 Antes de inserir:")
                            print(f"         - Valor atual do campo: '{nota_input.get_attribute('value')}'")
                            print(f"         - Nota a ser inserida: '{nota}'")
                            
                            # Limpa e insere a nota de forma BRUTA
                            nota_input.clear()
                            time.sleep(0.1)
                            
                            # Método BRUTO: insere caractere por caractere
                            print(f"      🔥 INSERÇÃO BRUTA: '{nota}'")
                            for char in nota:
                                nota_input.send_keys(char)
                                time.sleep(0.05)  # Pequena pausa entre caracteres
                            
                            time.sleep(0.1)
                            
                            # VERIFICAÇÃO FINAL: se não inseriu corretamente, força com JavaScript
                            valor_final = nota_input.get_attribute('value')
                            if valor_final != nota:
                                print(f"      ⚠️  Inserção falhou! Forçando com JavaScript...")
                                # Força a inserção usando JavaScript
                                driver.execute_script(f"arguments[0].value = '{nota}';", nota_input)
                                # Dispara evento de mudança
                                driver.execute_script("arguments[0].dispatchEvent(new Event('input', {{ bubbles: true }}));", nota_input)
                                time.sleep(0.1)
                                
                                valor_js = nota_input.get_attribute('value')
                                print(f"      🔥 JAVASCRIPT FORÇADO: '{valor_js}'")
                            else:
                                print(f"      ✅ Inserção bem-sucedida!")
                            
                            # Debug após inserir
                            valor_apos = nota_input.get_attribute('value')
                            print(f"      🔍 Após inserir:")
                            print(f"         - Valor do campo: '{valor_apos}'")
                            print(f"         - Nota esperada: '{nota}'")
                            print(f"         - Inserção bem-sucedida: {valor_apos == nota}")
                            
                            print(f"      ✅ Nota {nota} inserida na linha {i}")
                            return True
                        else:
                            print(f"      ❌ Input de nota não encontrado na linha {i}")
                            return False
                            
            except Exception as e:
                print(f"      ⚠️  Erro ao processar linha {i}: {e}")
                continue
                    
    except Exception as e:
        print(f"      💥 Erro ao processar tabela: {e}")
        import traceback
        print(f"      📋 Stack trace: {traceback.format_exc()}")
        return False
    
    print(f"      ❌ Matrícula {matricula} não encontrada em nenhuma linha")
    return False


# =========================================================
# MAIN
# =========================================================
def main():
    # 1) Lê planilha
    print("📊 Lendo planilha do Google Sheets...")
    ws = open_worksheet(SHEET_ID, WORKSHEET_NAME)
    records = read_rows_as_dicts(ws)
    print(f"✅ Encontradas {len(records)} linhas na planilha")

    # 2) Abre navegador e vai para a tela
    print("\n🌐 Abrindo navegador...")
    driver = make_driver(headless=False)
    driver.get(TARGET_URL)
    
    # 3) Faz login se necessário
    if LOGIN_NECESSARIO:
        print("🔐 Fazendo login...")
        do_login_if_needed(driver)
    
    # 4) Aguarda usuário estar pronto
    print("\n" + "="*60)
    print("📝 INSTRUÇÕES:")
    print("1. Navegue manualmente até a página de digitação de notas")
    print("2. Certifique-se de que está na tela correta")
    print("3. Quando estiver pronto para começar a digitação automática,")
    print("   pressione ENTER no terminal")
    print("="*60)
    
    input("\n⏳ Pressione ENTER quando estiver pronto para começar...")
    
    # 5) Verifica se está na página correta
    print("\n🔍 Verificando se está na página correta...")
    
    # Tenta encontrar elementos com seletores atuais
    elementos_encontrados = False
    max_tentativas = 3
    
    for tentativa in range(max_tentativas):
        try:
            if USE_IFRAME:
                switch_into_iframe_if_needed(driver)
                print("✅ Iframe carregado com sucesso")
            
            # Testa se consegue encontrar elementos básicos
            if USE_TABLE_STRATEGY:
                # Procura por linhas da tabela e inputs de nota
                rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                nota_inputs = driver.find_elements(By.CSS_SELECTOR, "input.nota-input")
                
                if rows and nota_inputs:
                    print(f"✅ Encontradas {len(rows)} linhas na tabela e {len(nota_inputs)} inputs de nota")
                    
                    # Mostra algumas linhas para debug
                    for i, row in enumerate(rows[:3]):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        if len(cells) >= 4:
                            matricula = cells[1].text.strip()
                            print(f"   📝 Linha {i}: Matrícula '{matricula}'")
                else:
                    raise Exception("Linhas da tabela ou inputs de nota não encontrados")
            else:
                wait_el(driver, **SELECTORS["search_input"])
                print("✅ Campo de busca encontrado")
                
            print("✅ Página carregada corretamente!")
            elementos_encontrados = True
            break
            
        except Exception as e:
            print(f"❌ Tentativa {tentativa + 1}/{max_tentativas}: Elementos não encontrados")
            
            if tentativa < max_tentativas - 1:
                print("\n🔧 AJUSTE OS SELETORES:")
                print("1. Verifique se está na página correta de digitação de notas")
                print("2. Inspecione os elementos da página (F12)")
                print("3. Ajuste os seletores no código se necessário")
                print("4. Pressione ENTER para tentar novamente...")
                input()
            else:
                print(f"\n💥 ERRO FINAL: Não foi possível encontrar os elementos necessários")
                print(f"   URL atual: {driver.current_url}")
                print(f"   Erro: {e}")
                print("\n🔧 SOLUÇÕES:")
                print("1. Verifique se está na página correta de digitação de notas")
                print("2. Os seletores CSS podem estar incorretos")
                print("3. A página pode ter mudado de estrutura")
                print("4. Pode ser necessário fazer login primeiro")
                
                # Permite continuar mesmo com erro
                print("\n⚠️  Deseja continuar mesmo com erro? (s/n): ", end="")
                continuar = input().lower().strip()
                if continuar != 's':
                    driver.quit()
                    return
                else:
                    print("⚠️  Continuando com elementos não encontrados...")
                    elementos_encontrados = False
                    break

    # 6) Processa todas as notas automaticamente
    if elementos_encontrados:
        print(f"\n🚀 Iniciando digitação automática de {len(records)} notas...")
        print("   (Pressione Ctrl+C para parar a qualquer momento)")
        print("-" * 60)
        
        not_found = []
        total_ok = 0

        for i, row in enumerate(records, start=1):
            matricula = str(row[HEADER_MATRICULA]).strip()
            nota_original = row[HEADER_NOTA]
            nota = normalizar_nota(row[HEADER_NOTA])

            if nota == "":
                print(f"[{i:3d}] Matrícula {matricula}: sem nota (TOTAL vazio) — pulando.")
                continue

            print(f"[{i:3d}] 🔍 Processando: {matricula}")
            print(f"      📊 Tipo da nota: {type(nota_original)}")
            print(f"      📝 Nota original: '{nota_original}'")
            print(f"      🔧 Nota normalizada: '{nota}'")
            print(f"      📏 Comprimento: {len(nota)}")
            print(f"      🔍 Contém vírgula: {',' in nota}")
            print(f"      🔍 Contém ponto: {'.' in nota}")

            try:
                if USE_TABLE_STRATEGY:
                    ok = set_grade_in_table_row(driver, matricula, nota)
                    if not ok:
                        not_found.append(matricula)
                        print(f"[{i:3d}] ❌ NÃO ENCONTRADA na tabela: {matricula}")
                    else:
                        total_ok += 1
                        print(f"[{i:3d}] ✅ OK {matricula} -> {nota}")
                else:
                    # Estratégia de busca por matrícula
                    set_grade_using_search(driver, matricula, nota)
                    total_ok += 1
                    print(f"[{i:3d}] ✅ OK {matricula} -> {nota} (via busca)")

            except Exception as e:
                print(f"[{i:3d}] 💥 ERRO para {matricula}: {e}")

        # 7) Resultado final
        switch_out_iframe(driver)
        print("\n" + "="*60)
        print("📊 RESUMO FINAL:")
        print(f"   ✅ Notas digitadas com sucesso: {total_ok}")
        print(f"   ❌ Matrículas não encontradas: {len(not_found)}")
        
        if not_found:
            print("\n🔍 Matrículas não encontradas (confira seletores/iframe/paginação):")
            for m in not_found:
                print("   •", m)
    else:
        print(f"\n⚠️  MODO MANUAL: Elementos não encontrados automaticamente")
        print(f"   📋 Você tem {len(records)} notas para digitar manualmente:")
        print("-" * 60)
        
        for i, row in enumerate(records, start=1):
            matricula = str(row[HEADER_MATRICULA]).strip()
            nota = normalizar_nota(row[HEADER_NOTA])
            if nota != "":
                print(f"[{i:3d}] {matricula} -> {nota}")
        
        print("\n💡 DICA: Use Ctrl+F na página para buscar as matrículas e digite as notas manualmente")
        
        # 7) Resultado final (modo manual)
        print("\n" + "="*60)
        print("📊 MODO MANUAL ATIVADO:")
        print(f"   📝 Notas para digitar: {len([r for r in records if normalizar_nota(r[HEADER_NOTA]) != ''])}")
        print("   🔍 Digite as notas manualmente na página")

    # 8) Aguarda usuário verificar e enviar manualmente
    print("\n" + "="*60)
    print("🔍 VERIFICAÇÃO MANUAL:")
    print("1. Verifique se todas as notas foram digitadas corretamente")
    print("2. Envie/salve as notas manualmente no sistema")
    print("3. Quando terminar, volte ao terminal e pressione ENTER")
    print("="*60)
    
    input("\n⏳ Pressione ENTER quando terminar de verificar e enviar as notas...")
    
    # 9) Confirmação final
    print("\n✅ Processo finalizado!")
    print("   O navegador permanecerá aberto para você.")
    print("   Feche-o manualmente quando desejar.")
    
    # 10) Solicita novo SHEET_ID para próxima turma
    print("\n" + "="*60)
    print("🔄 PRÓXIMA TURMA:")
    print("1. Feche esta janela do Chrome")
    print("2. Navegue manualmente para a página da próxima turma")
    print("3. Cole o novo SHEET_ID da próxima turma")
    print("="*60)
    
    novo_sheet_id = input("\n📋 Cole o SHEET_ID da próxima turma (ou pressione ENTER para sair): ").strip()
    
    if novo_sheet_id:
        print(f"\n🔄 Atualizando SHEET_ID para: {novo_sheet_id}")
        print("   Execute o programa novamente com o novo ID!")
    else:
        print("\n👋 Programa finalizado!")
    
    # 11) Mantém o navegador aberto (não fecha automaticamente)
    print("\n🌐 Navegador mantido aberto.")
    print("   Para fechar, feche a janela do Chrome ou pressione Ctrl+C no terminal.")
    
    try:
        # Mantém o programa rodando até o usuário interromper
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n👋 Encerrando programa...")
        driver.quit()
        print("✅ Navegador fechado. Programa finalizado!")


if __name__ == "__main__":
    # Dependências:
    # pip install gspread google-auth selenium webdriver-manager
    main()