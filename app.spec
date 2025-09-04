block_cipher = None  # 用于加密 Python 文件的密钥，通常不需要设置

a = Analysis(['main.py'],  # 指定要打包的主脚本文件
             pathex=['H:/chaos项目资料/特征解析工具汇编/PHM2.0打包'],  # 项目路径，PyInstaller 查找文件的路径
             binaries=[],  # 外部二进制文件，不需要添加任何二进制文件
             datas=[],  # 附加的数据文件，不需要添加任何数据文件
             hiddenimports=[],  # 隐式导入的模块列表，通常不需要添加任何模块
             hookspath=[],  # 自定义 hook 文件的路径，通常保持为空
             runtime_hooks=[],  # 自定义的运行时 hooks，通常保持为空
             excludes=[],  # 排除的模块列表，通常不需要排除任何模块
             win_no_prefer_redirects=False,  # 是否在 Windows 上优先使用重定向，通常保持为 False
             win_private_assemblies=False,  # 是否将 DLL 文件标记为私有，通常保持为 False
             cipher=block_cipher,  # 加密模块，通常设置为 None
             noarchive=False)  # 是否不将所有文件归档，通常保持为 False

pyz = PYZ(a.pure, a.zipped_data,  # 创建 PYZ 文件，包含纯 Python 文件和压缩数据
          cipher=block_cipher)  # 使用的加密密钥，通常设置为 None

exe = EXE(pyz,  # 创建 EXE 文件
          a.scripts,  # 要打包的脚本列表
          exclude_binaries=True,  # 排除二进制文件
          name='特征解析工具V3.0（内部版本）',  # 生成的可执行文件名称
          debug=False,  # 是否启用调试模式，通常保持为 False
          bootloader_ignore_signals=False,  # 是否忽略启动程序的信号，通常保持为 False
          strip=False,  # 是否移除调试信息，通常保持为 False
          upx=True,  # 是否使用 UPX 压缩，通常保持为 True
          console=False,  # 是否启用控制台，True 表示启用，False 表示禁用
          icon='images\\logo.ico')  # 指定应用程序的图标文件路径

coll = COLLECT(exe,  # 收集生成的 EXE 文件
               a.binaries,  # 包含的二进制文件
               a.zipfiles,  # 包含的 ZIP 文件
               a.datas,  # 包含的数据文件
               strip=False,  # 是否移除调试信息，通常保持为 False
               upx=True,  # 是否使用 UPX 压缩，通常保持为 True
               name='dist/特征解析工具V3.0（内部版本）')  # 输出文件夹和生成的可执行文件名称
