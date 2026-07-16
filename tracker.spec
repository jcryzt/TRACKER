# tracker.spec - PyInstaller spec para gerar tracker.exe (onefile)
import os
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

hidden_imports = [
    'msal', 'msal.application', 'msal.authority', 'msal.token_cache',
    'msal.oauth2cli', 'msal.oauth2cli.oidc',
    'cryptography', 'cryptography.hazmat.primitives',
    'cryptography.hazmat.backends.openssl',
    'flask', 'flask.templating',
    'flask_limiter', 'flask_limiter.util', 'flask_limiter.wrappers',
    'limits', 'limits.storage', 'limits.strategies',
    'requests', 'requests.adapters',
    'urllib3', 'certifi',
    'jinja2', 'jinja2.ext',
    'werkzeug', 'werkzeug.routing', 'werkzeug.serving',
    'ctypes', 'ctypes.wintypes',
]
hidden_imports += collect_submodules('msal')
hidden_imports += collect_submodules('limits')

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'pandas', 'PIL', 'test', 'unittest'],
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
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='static/favicon.ico',
)
