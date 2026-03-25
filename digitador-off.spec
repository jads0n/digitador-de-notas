# -*- mode: python ; coding: utf-8 -*-
# Arquivo de configuração do PyInstaller — funciona em Windows e macOS
# Execute: pyinstaller digitador-off.spec --clean --noconfirm

import sys
from PyInstaller.utils.hooks import collect_data_files

# Coleta todos os arquivos de dados do customtkinter (temas, fontes, imagens)
ctk_datas = collect_data_files("customtkinter")

block_cipher = None

a = Analysis(
    ['digitador-off.py'],
    pathex=[],
    binaries=[],
    datas=ctk_datas,
    hiddenimports=[
        # customtkinter / tkinter
        'customtkinter',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        # Selenium — módulos principais
        'selenium',
        'selenium.webdriver',
        # Chrome
        'selenium.webdriver.chrome',
        'selenium.webdriver.chrome.options',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.chrome.webdriver',
        # Firefox (opcional, evita erros de importação interna)
        'selenium.webdriver.firefox',
        'selenium.webdriver.firefox.options',
        'selenium.webdriver.firefox.service',
        # Remote / Common
        'selenium.webdriver.remote',
        'selenium.webdriver.remote.webdriver',
        'selenium.webdriver.remote.command',
        'selenium.webdriver.remote.remote_connection',
        'selenium.webdriver.common',
        'selenium.webdriver.common.by',
        'selenium.webdriver.common.keys',
        'selenium.webdriver.common.action_chains',
        'selenium.webdriver.common.desired_capabilities',
        'selenium.webdriver.common.options',
        'selenium.webdriver.common.service',
        'selenium.webdriver.common.utils',
        # Support
        'selenium.webdriver.support',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.wait',
        'selenium.webdriver.support.expected_conditions',
        # webdriver_manager
        'webdriver_manager',
        'webdriver_manager.chrome',
        'webdriver_manager.core',
        'webdriver_manager.core.driver',
        'webdriver_manager.core.manager',
        'webdriver_manager.core.utils',
        # Dependências de rede / certificados
        'packaging',
        'packaging.version',
        'packaging.specifiers',
        'packaging.requirements',
        'certifi',
        'urllib3',
        'urllib3.util',
        'urllib3.util.retry',
        'requests',
        'requests.adapters',
        'requests.auth',
        'requests.packages',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DigitadorDeNotas',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # False = sem janela de terminal preta
    disable_windowed_traceback=False,
    argv_emulation=False,   # Necessário para macOS
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',      # Windows: descomente e coloque seu ícone .ico
    # icon='icon.icns',     # macOS:   descomente e coloque seu ícone .icns
)

# No macOS, cria também um pacote .app (opcional, mas mais elegante)
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='DigitadorDeNotas.app',
        icon=None,           # Coloque 'icon.icns' se tiver um ícone
        bundle_identifier='com.digitador.notas',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleDisplayName': 'Digitador de Notas',
            'CFBundleVersion': '1.0.0',
        },
    )
