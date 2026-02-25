APP_NAME = 'MaxDistanceCalculator_v1.31'

# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['D:/Python_workspace/max_distance_calculator/MaxDistanceCalculator.py'],
    pathex=[],
    binaries=[],
    datas=[('D:/Python_workspace/max_distance_calculator/app_icon.ico', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,         # 추가: 바이너리 포함
    a.datas,            # 추가: 데이터 파일 포함
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None, # 추가: 임시 디렉토리 설정
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['D:/Python_workspace/max_distance_calculator/app_icon.ico'],
)

# COLL 섹션은 One File 빌드 시 필요 없으므로 삭제하거나 주석 처리합니다.
