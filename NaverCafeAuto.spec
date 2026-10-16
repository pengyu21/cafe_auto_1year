# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules

# 크롬 '비밀번호 저장' 팝업 클릭용 (윈도우 UI 자동화)
uia_datas, uia_binaries, uia_hiddenimports = collect_all('uiautomation')


a = Analysis(
    ['gui_cafeauto.py'],
    pathex=[],
    binaries=uia_binaries,
    datas=uia_datas,
    hiddenimports=[
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.chrome.webdriver',
        'selenium.webdriver.common.action_chains',
        'selenium.webdriver.common.keys',
        'selenium.webdriver.common.by',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.expected_conditions',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.chrome.options',
        'webdriver_manager'
    ] + uia_hiddenimports + collect_submodules('comtypes'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='NaverCafeAuto',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
