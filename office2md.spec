# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path

block_cipher = None

# ImageMagick binary path
imagemagick_path = r'C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe'

# Magika models path
import sys
if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
    # Virtual environment
    magika_path = os.path.join(sys.prefix, 'Lib', 'site-packages', 'magika')
else:
    # System Python
    magika_path = os.path.join(sys.base_prefix, 'Lib', 'site-packages', 'magika')

# Data files to include
datas = []

# Add magika models and config
if os.path.exists(magika_path):
    models_path = os.path.join(magika_path, 'models')
    if os.path.exists(models_path):
        datas.append((models_path, 'magika/models'))

    config_path = os.path.join(magika_path, 'config')
    if os.path.exists(config_path):
        datas.append((config_path, 'magika/config'))

# Binaries to include
binaries = []
if os.path.exists(imagemagick_path):
    binaries.append((imagemagick_path, 'bin'))

# Hidden imports - packages that PyInstaller might miss
hiddenimports = [
    'markitdown',
    'mammoth',
    'mammoth.docx',
    'mammoth.docx.body_xml',
    'mammoth.docx.office_xml',
    'docxlatex',
    'docxlatex.parser',
    'PIL',
    'PIL.Image',
    'markdownify',
    'bs4',
    'lxml',
    'openai',
    'base64',
    'magika',
    'magika.magika',
    'onnxruntime',
]

a = Analysis(
    ['packages/markitdown/src/markitdown/office2md.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['transformers', 'torch', 'tensorflow', 'datasets'],
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
    name='office2md',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Console application
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Can add icon file here
)
