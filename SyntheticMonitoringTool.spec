# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_dynamic_libs
from PyInstaller.utils.hooks import collect_submodules

SPEC_PATH = globals().get("__file__", os.path.join(os.getcwd(), "SyntheticMonitoringTool.spec"))
BASE_DIR = os.path.dirname(os.path.abspath(SPEC_PATH))
APP_ICON = os.path.join(BASE_DIR, "app_icon.ico")

datas = [
    (os.path.join(BASE_DIR, 'app.py'), '.'),
    (os.path.join(BASE_DIR, 'app_icon.ico'), '.'),
    (os.path.join(BASE_DIR, 'pw-browsers'), 'pw-browsers'),
]
binaries = []
hiddenimports = [
    'win32timezone',
    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
    'httpx',
    'bs4',
]
datas += collect_data_files('playwright')
binaries += collect_dynamic_libs('playwright')
hiddenimports += collect_submodules('playwright')

def _keep_toc_item(item):
    text = "|".join(str(x) for x in item) if isinstance(item, (tuple, list)) else str(item)
    text = text.lower()
    # Keep bundled browser binaries under project pw-browsers.
    if "pw-browsers" in text:
        return True
    blocked_keywords = [
        ".local-browsers",
        "pyside6\\translations\\",
        "pyside6/translations/",
        "pyside6\\plugins\\generic\\",
        "pyside6/plugins/generic/",
        "pyside6\\plugins\\networkinformation\\",
        "pyside6/plugins/networkinformation/",
        "pyside6\\plugins\\platforminputcontexts\\",
        "pyside6/plugins/platforminputcontexts/",
        "pyside6\\plugins\\platforms\\qdirect2d.dll",
        "pyside6/plugins/platforms/qdirect2d.dll",
        "pyside6\\plugins\\platforms\\qminimal.dll",
        "pyside6/plugins/platforms/qminimal.dll",
        "pyside6\\plugins\\platforms\\qoffscreen.dll",
        "pyside6/plugins/platforms/qoffscreen.dll",
        "pyside6\\plugins\\tls\\",
        "pyside6/plugins/tls/",
        "pyside6\\qt6pdf.dll",
        "pyside6/qt6pdf.dll",
        "pyside6\\qt6qml.dll",
        "pyside6/qt6qml.dll",
        "pyside6\\qt6qmlmeta.dll",
        "pyside6/qt6qmlmeta.dll",
        "pyside6\\qt6qmlmodels.dll",
        "pyside6/qt6qmlmodels.dll",
        "pyside6\\qt6qmlworkerscript.dll",
        "pyside6/qt6qmlworkerscript.dll",
        "pyside6\\qt6quick.dll",
        "pyside6/qt6quick.dll",
        "pyside6\\qt6virtualkeyboard.dll",
        "pyside6/qt6virtualkeyboard.dll",
        "pyside6\\qt6svg.dll",
        "pyside6/qt6svg.dll",
        "pyside6\\plugins\\iconengines\\qsvgicon.dll",
        "pyside6/plugins/iconengines/qsvgicon.dll",
        "pyside6\\plugins\\imageformats\\qsvg.dll",
        "pyside6/plugins/imageformats/qsvg.dll",
        "pyside6\\plugins\\imageformats\\qpdf.dll",
        "pyside6/plugins/imageformats/qpdf.dll",
    ]
    return not any(keyword in text for keyword in blocked_keywords)


# First-pass filtering before Analysis.
datas = [item for item in datas if _keep_toc_item(item)]
binaries = [item for item in binaries if _keep_toc_item(item)]


a = Analysis(
    ['run_tool.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'streamlit',
        'streamlit.web',
        'streamlit.external',
        'langchain',
        'matplotlib',
        'numpy.tests',
        'pandas.tests',
    ],
    noarchive=False,
    optimize=0,
)

# Second-pass filtering after Analysis; PyInstaller hooks may re-add
# Playwright internal browser cache entries that exceed Windows path limits.
a.datas = [item for item in a.datas if _keep_toc_item(item)]
a.binaries = [item for item in a.binaries if _keep_toc_item(item)]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SyntheticMonitoringTool',
    icon=APP_ICON,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
    upx=False,
    upx_exclude=[],
    name='SyntheticMonitoringTool',
)
