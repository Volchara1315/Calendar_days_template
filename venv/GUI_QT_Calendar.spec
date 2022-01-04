# -*- mode: python -*-

block_cipher = None


a = Analysis(['GUI_QT_Calendar.py'],
             pathex=['D:\\Python\\calendar_days_template\\venv'],
             binaries=[],
             datas=[],
             hiddenimports=['xlsxwriter'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='GUI_QT_Calendar',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='Treetog-Junior-Document-excel.ico')
