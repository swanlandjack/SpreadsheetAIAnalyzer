# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for Excel AI Analyzer backend
#
# Usage (from the backend/ directory):
#   pyinstaller server.spec
#
# Output:
#   backend/dist/server/server        (Mac/Linux)
#   backend/dist/server/server.exe    (Windows)
#
# This compiles server.py + all Python dependencies into a standalone
# directory that runs without Python being installed on the user's machine.

import sys
import os

block_cipher = None

OUT_NAME = 'server-win' if sys.platform.startswith('win') else 'server-mac'

a = Analysis(
    ['server.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # Flask internals that PyInstaller sometimes misses
        'flask',
        'flask_cors',
        'werkzeug',
        'werkzeug.routing',
        'werkzeug.exceptions',
        'werkzeug.serving',
        'werkzeug.middleware',
        'werkzeug.middleware.proxy_fix',
        'werkzeug.debug',
        'jinja2',
        'jinja2.ext',
        'click',
        'eparse',
        'eparse.core',
        'eparse.interfaces',
        'peewee',
        'lxml',
        'lxml.etree',

        # Pandas engines
        'pandas',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.offsets',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.interval',
        'pandas._libs.hashtable',
        'pandas._libs.lib',
        'pandas._libs.missing',
        'pandas._libs.reduction',
        'pandas._libs.reshape',
        'pandas._libs.sparse',
        'pandas.io.formats.style',
        'pandas.io.formats.excel',

        # Excel engines
        'openpyxl',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'openpyxl.reader.excel',
        'openpyxl.styles',
        'xlrd',
        'xlrd.biffh',
        'xlrd.book',
        'xlrd.compdoc',
        'xlrd.formatting',

        # Numpy
        'numpy',
        'numpy.core._multiarray_umath',
        'numpy.core._operand_flag_tests',
        'numpy.core._rational_helpers',
        'numpy.core._struct_ufunc_tests',
        'numpy.core._umath_tests',

        # Requests / networking
        'requests',
        'urllib3',
        'certifi',
        'charset_normalizer',
        'idna',

        # Standard library
        'io',
        'json',
        'os',
        're',
        'traceback',
        'threading',
        'socket',
        'ssl',
        'email',
        'email.mime',
        'email.mime.text',
        'http',
        'http.server',
        'http.client',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Things we definitely do not need — reduces binary size
        'tkinter',
        'psycopg2',
        'psycopg2-binary',
        'matplotlib',
        'scipy',
        'sklearn',
        'PIL',
        'cv2',
        'IPython',
        'notebook',
        'jupyter',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'wx',
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
    [],
    exclude_binaries=True,
    name='server',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,          # compress with UPX if available (reduces size ~30%)
    console=True,      # keep console=True so Flask logs are visible during dev/debug
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,  # None = native arch; set 'arm64' or 'x86_64' to cross-compile
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
    name=OUT_NAME,
)
