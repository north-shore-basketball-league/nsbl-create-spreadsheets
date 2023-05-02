# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['./src/updater/__init__.py','./src/updater/version.py'],
    pathex=['./src/updater'],
    binaries=[],
    hiddenimports=["xlwings", "thefuzz.fuzz", "dateutil.parser", "re", "selenium", "webdriver_manager", "webdriver_manager.chrome", "webdriver_manager.chrome.ChromeDriverManager", "selenium", "selenium.webdriver", "selenium.webdriver.chrome", "selenium.webdriver.support", "selenium.webdriver.common", "selenium.webdriver.chrome.service", "selenium.webdriver.support.wait", "selenium.webdriver.common.by", "pandas", "json", "pathlib"],
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
    [],
    exclude_binaries=True,
    name='Extract NSBL Games',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Extract NSBL Games',
)
