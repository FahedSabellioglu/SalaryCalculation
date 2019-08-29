# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['gui.py'],
             pathex=['C:\\Users\\Fahed Sabellioglu\\PycharmProjects\\Ankara',
		     r'C:\Users\Fahed Sabellioglu\PycharmProjects\Ankara\venv\Lib\site-packages'],
             binaries=[],
             datas=[],
             hiddenimports=["pandas","numpy","openpyxl","time","datetime",
			"numpy.random.common","numpy.random.bounded_integers",
			"numpy.random.entropy"],
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
          a.binaries+[("rocket+1.ico",r"C:\Users\Fahed Sabellioglu\PycharmProjects\Ankara\rocket+1.ico","DATA")],
          a.zipfiles,
          a.datas,
          [],
          name='JKE',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,
icon="rocket+1.ico" )
