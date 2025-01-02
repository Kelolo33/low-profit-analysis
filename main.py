import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'

import sys
from gui import run_gui
from analyze_data import analyze_excel_data
import time

def check_dependencies():
    required_packages = ['pandas', 'openpyxl', 'xlrd']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"缺少必要的依赖包：{', '.join(missing_packages)}，正在安装...")
        try:
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing_packages)
            print("依赖包安装完成")
        except subprocess.CalledProcessError as e:
            print(f"安装依赖包时出错：{str(e)}")
            sys.exit(1)

def main():
    check_dependencies()
    
    # 运行 GUI
    input_file, output_file, subscription_file = run_gui()

if __name__ == "__main__":
    main()
