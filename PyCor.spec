# -*- mode: python -*-

block_cipher = None

a = Analysis(['src\\__init__.py'],
             pathex=['src'],
             binaries=[],
             datas=[
                ('src/config.example.py', '.'),
                ('src/cacert.pem','.')
             ],
             hiddenimports=[
                'ipaddress',
                'sentry_sdk.integrations.argv',
                'sentry_sdk.integrations.atexit',
                'sentry_sdk.integrations.dedupe',
                'sentry_sdk.integrations.excepthook',
                'sentry_sdk.integrations.logging',
                'sentry_sdk.integrations.modules',
                'sentry_sdk.integrations.stdlib',
                'sentry_sdk.integrations.threading',
                'win32timezone'
              ],
             hookspath=[],
             runtime_hooks=[],
             excludes=['config'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='PyCor',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
          runtime_tmpdir=None,
          console=True,
          version='version.txt')

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='PyCor')