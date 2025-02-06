import PyInstaller.__main__
import os
import sys
import site

def find_tkdnd_path():
    # 获取所有site-packages路径
    site_packages = site.getsitepackages()
    # 在所有路径中查找tkinterdnd2
    for path in site_packages:
        tkdnd_path = os.path.join(path, 'tkinterdnd2')
        if os.path.exists(tkdnd_path):
            return tkdnd_path
    return None

def build():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 获取tkinterdnd2模块的路径
    tkdnd_path = find_tkdnd_path()
    if not tkdnd_path:
        print("Error: Cannot find tkinterdnd2 module")
        return
    
    # 设置图标路径
    icon_path = os.path.join(current_dir, 'icon.png')
    
    PyInstaller.__main__.run([
        'main.py',  # 主程序文件
        '--name=Excel合并工具',  # 生成的exe文件名
        '--windowed',  # 使用GUI模式，不显示控制台
        '--onefile',  # 打包成单个exe文件
        f'--icon={icon_path}',  # 设置图标（如果需要）
        '--clean',  # 清理临时文件
        '--noconfirm',  # 不确认覆盖
        f'--add-data={tkdnd_path};tkinterdnd2',  # 添加tkinterdnd2模块
    ])

if __name__ == "__main__":
    build() 