# -*- mode: python ; coding: utf-8 -*-
# tracker.spec — PyInstaller spec para gerar tracker.exe (onefile)
#
# Como usar:
#   pip install pyinstaller
#   pyinstaller tracker.spec
#
# O executável será gerado em: dist/tracker.exe
# Coloque ao lado do tracker.exe:
#   - config.py        (configurações — NÃO incluído no bundle)
#   - seeds/           (criado automaticamente ao iniciar)
#
# tracker_data.json e .sp_token_cache.bin são criados automaticamente.

import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# Hidden imports necessários para msal, flask_limiter e dependências
hidden_imports = [
    # MSAL / autenticação Microsoft
    'msal',
    'msal.application',
    'msal.authority',
    'msal.token_cache',
    'msal.oauth2cli',
    'msal.oauth2cli.oidc',
    # Criptografia (usada pelo msal)
    'cryptography',
    'cryptography.hazmat.primitives',
    'cryptography.hazmat.backends.openssl',
    # Flask e extensões
    'flask',
    'flask.templating',
    'flask_limiter',
    'flask_limiter.util',
    'flask_limiter.wrappers',
    'limits',
    'limits.storage',
    'limits.strategies',
    # Requests
    'requests',
    'requests.adapters',
    'urllib3',
    'certifi',
    # Jinja2 / Werkzeug
    'jinja2',
    'jinja2.ext',
    'werkzeug',
    'werkzeug.routing',
    'werkzeug.serving',
    # Windows
    'ctypes',
    'ctypes.wintypes',
    'win32api',
    'win32con',
]

# Coleta todos os submodulos do msal e limits
hidden_imports += collect_submodules('msal')
hidden_imports += collect_submodules('limits')

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # Empacota templates e static dentro do bundle
        ('templates', 'templates'),
        ('static',    'static'),
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'PIL',
        'test',
        'unittest',
    ],
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
    name='tracker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,           # UPX pode causar falsos positivos em antivírus
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,        # Mostra console (útil para ver erros/logs)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='static/favicon.ico',
)
