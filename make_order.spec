# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['make_order.py'],
             pathex=['C:\\Users\\GP73\\PycharmProjects\\Dispatch_helper\\virtualenv_dispatch_helper\\merge_dispatch_helper'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
    
b = Analysis(['make_dispatch_order.py'],
             pathex=['C:\\Users\\GP73\\PycharmProjects\\Dispatch_helper\\virtualenv_dispatch_helper\\merge_dispatch_helper'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

c = Analysis(['delete_files.py'],
             pathex=['C:\\Users\\GP73\\PycharmProjects\\Dispatch_helper\\virtualenv_dispatch_helper\\merge_dispatch_helper'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)    
  
  
MERGE( (a, 'make_order', 'make_order'), (b, 'make_dispatch_order', 'make_dispatch_order'), (c, 'delete_files', 'delete_files') )    
    
    
make_order_pyz = PYZ(a.pure, 
         a.zipped_data,
      cipher=block_cipher)
      
make_order_exe = EXE(make_order_pyz,
      a.scripts,
      [],
      exclude_binaries=True,
      name='make_order',
      debug=True,
      bootloader_ignore_signals=False,
      strip=False,
      upx=True,
      console=False )      

make_dispatch_order_pyz = PYZ(b.pure, 
            b.zipped_data,
         cipher=block_cipher)

make_dispatch_order_exe = EXE(make_dispatch_order_pyz,
            b.scripts,
            [],
            exclude_binaries=True,
            name='make_dispatch_order',
            debug=True,
            bootloader_ignore_signals=False,
            strip=False,
            upx=True,
            console=False )
    
delete_files_pyz = PYZ(c.pure, 
        c.zipped_data,
        cipher=block_cipher)

delete_files_exe = EXE(delete_files_pyz,
       c.scripts,
       [],
       exclude_binaries=True,
       name='delete_files',
       debug=True,
       bootloader_ignore_signals=False,
       strip=False,
       upx=True,
       console=False )


coll = COLLECT(make_order_exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               make_dispatch_order_exe,
               b.binaries,
               b.zipfiles,
               b.datas,			   
               delete_files_exe,
               c.binaries,
               c.zipfiles,
               c.datas,			   
               strip=False,
               upx=True,
               upx_exclude=[],
               name='dispatch_helper')
			   
