import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'

import sys
from gui import run_gui
from analyze_data import analyze_excel_data
import time

def check_dependencies():
    try:
        import pandas
        import openpyxl
        import xlrd
    except ImportError as e:
        print("缺少必要的依赖包，正在安装...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl", "xlrd>=2.0.1"])
        print("依赖包安装完成")

def main():
    check_dependencies()
    
    # 运行 GUI
    input_file, output_file, subscription_file = run_gui()

if __name__ == "__main__":
    main()