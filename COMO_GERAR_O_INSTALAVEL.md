# 🤖 Digitador de Notas — Como Gerar o Instalável para Windows

## Pré-requisitos

1. **Python 3.10 ou superior** instalado no seu computador Windows  
   Download: https://www.python.org/downloads/  
   > ⚠️ **IMPORTANTE**: Na tela de instalação, marque a opção **"Add Python to PATH"**

2. **Google Chrome** instalado (necessário tanto para gerar quanto para usar o programa)

---

## Método Rápido (Recomendado)

1. Copie toda a pasta do projeto para o seu computador **Windows**
2. Dentro da pasta, dê um duplo-clique em **`build_windows.bat`**
3. Aguarde (pode levar de 2 a 5 minutos na primeira vez)
4. Quando terminar, acesse a pasta **`dist`** — lá estará o arquivo **`DigitadorDeNotas.exe`**

---

## Método Manual (Passo a Passo)

Abra o **Prompt de Comando (CMD)** ou **PowerShell** na pasta do projeto e execute:

```bat
:: Instalar as dependências
pip install customtkinter selenium webdriver-manager pyinstaller

:: Gerar o executável
pyinstaller digitador-off.spec --clean --noconfirm
```

---

## Usando o Executável Gerado

O arquivo `dist\DigitadorDeNotas.exe`:

- **NÃO precisa** do Python instalado para rodar
- Pode ser copiado para qualquer computador Windows
- **PRECISA** do **Google Chrome** instalado no computador de destino
- O **ChromeDriver** é baixado automaticamente na primeira execução (requer internet)

---

## Formato do Arquivo CSV

O arquivo CSV deve ter **exatamente** estas colunas:

| MATRICULA | TOTAL |
|-----------|-------|
| 123456    | 8,5   |
| 789012    | 7,0   |

- O separador pode ser `;` ou `,`
- A codificação pode ser UTF-8 ou ISO-8859-1 (gerado pelo Excel BR)

---

## Solução de Problemas

| Problema | Solução |
|----------|---------|
| `'python' não é reconhecido` | Python não está no PATH. Reinstale marcando "Add to PATH" |
| `ModuleNotFoundError` | Execute `pip install -r requirements.txt` |
| `ChromeDriver error` | Verifique se o Chrome está instalado e atualizado |
| Tela preta aparece | Normal no build, fecha sozinha quando o app abre |
| Antivírus bloqueia | Adicione o `.exe` como exceção — é falso positivo comum em apps PyInstaller |

---

## Estrutura dos Arquivos

```
Digitador de notas/
├── digitador-off.py          # Código-fonte principal
├── digitador-off.spec        # Configuração do PyInstaller
├── build_windows.bat         # Script de build automático
├── requirements.txt          # Lista de dependências
└── dist/
    └── DigitadorDeNotas.exe  # ← Executável final (gerado após o build)
```
