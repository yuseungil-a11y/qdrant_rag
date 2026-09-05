# -*- mode: python ; coding: utf-8 -*-
import glob
import os
import sys
import sysconfig
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []
tmp_ret = collect_all('mcp')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('httpx2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('httpcore2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('anyio')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('hwp5')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('certifi')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pptx')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('docx')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pytesseract')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PIL')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('truststore')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('mwclient')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('lxml')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('bs4')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
# pydantic_core는 Rust로 컴파일된 확장 모듈이라, PyInstaller의 자동 의존성 분석이 이걸
# 놓치는 경우가 있다 - 이 경우 개발 PC(모듈이 이미 설치돼 있어 어떤 경로로든 로드됨)에서는
# 멀쩡히 동작하는 것처럼 보여도, 다른 PC(exe만 실행)에서는
# "ModuleNotFoundError: No module named 'pydantic_core._pydantic_core'"로 시작부터 죽는다
# (2026-09-05 실사용 중 발견 - AMD PC에서 재현, 원인은 CPU 제조사와 무관하고 순수 번들링 누락).
tmp_ret = collect_all('pydantic')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pydantic_core')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# pywintypes/pythoncom(pywin32)도 pydantic_core와 같은 이유로 자동 분석에서 누락된다.
# pywin32는 컴파일된 DLL(pywintypesNNN.dll/pythoncomNNN.dll)을 site-packages\pywin32_system32\
# 에 따로 두고, pywintypes.py가 실행 시점에 그 폴더를 직접 찾아서 로드하는 특이한 구조라
# 일반적인 import 분석/기존 pyinstaller-hooks-contrib 훅만으로는 이 DLL이 안 실린 채
# 빌드되는 경우가 있다 - 개발 PC에서는 pywin32가 시스템에 설치돼 있어(pip install 시
# System32 등에도 DLL이 등록됨) 이 누락이 가려져 정상 동작하는 것처럼 보였지만, 파이썬이
# 아예 없는 배포 대상 PC(요구사항: "파이썬이 설치 안되는 유저들에게도 배포")에서는
# "Module 'pywintypes' isn't in frozen sys.path"로 시작부터 죽는다(2026-09-05 실사용 중
# 발견, AMD PC였으나 원인은 CPU와 무관). pywintypes.py의 검색 경로와 정확히 같은 폴더명
# (pywin32_system32)으로 DLL을 직접 넣어서 훅 동작 여부와 무관하게 확실히 포함되게 한다.
_site_packages = sysconfig.get_paths()["purelib"]
_pywin32_system32_dir = os.path.join(_site_packages, "pywin32_system32")
for _dll in glob.glob(os.path.join(_pywin32_system32_dir, "*.dll")):
    binaries.append((_dll, "pywin32_system32"))
tmp_ret = collect_all('pythoncom')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pywintypes')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('win32com')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'pandas', 'scipy', 'PySide6', 'shiboken6'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

if sys.platform == 'darwin':
    # onefile + macOS .app(windowed) 조합은 PyInstaller 자체가 "macOS 보안 체계와 충돌한다"고
    # 경고하는 조합이다 (실행할 때마다 임시 폴더에 압축 해제하는 게 Gatekeeper/코드사인 검증과 겹쳐
    # 첫 더블클릭은 조용히 실패하고 두 번째부터 실행되는 증상 발생). macOS에서는 onedir로 빌드해서 회피.
    # 주의: pyinstaller를 spec 파일 없이(예: pyinstaller gui.py ...) 실행하면 이 spec이 통째로
    # 재생성되며 이 분기가 사라진다 - 반드시 `pyinstaller qdrant_register_gui.spec`으로만 빌드할 것.
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name='utinfo_vdr',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name='utinfo_vdr',
    )
    app = BUNDLE(
        coll,
        name='utinfo_vdr.app',
        icon=None,
        bundle_identifier='co.kr.utinfo.qdrant-register-gui',
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name='utinfo_vdr',
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
        icon='icon.ico',
    )
