import tkinter as tk
from tkinter import filedialog, messagebox
import analyze_data
import threading  # 导入 threading 模块
import pandas as pd
import time

class DataAnalysisGUI:
    def __init__(self, master):
        self.master = master
        master.title("数据分析工具")
        master.geometry("400x320")  # 稍微增加高度
        
        # 添加运行标志和窗口关闭处理
        self.is_running = False
        self.start_time = None
        self.processing_time = 0
        master.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 设置窗口背景色为浅色
        master.configure(bg='#F0F0F0')  # 使用浅灰色背景

        self.input_file = ""
        self.output_file = ""
        self.subscription_file = ""

        button_width = 20
        button_height = 1  # 减小按钮高度使其更符合 Mac 风格
        label_width = 12

        # 创建主框架并设置背景色
        main_frame = tk.Frame(master, bg='#F0F0F0')
        main_frame.pack(expand=True, fill=tk.BOTH, padx=30, pady=30)  # 增加边距

        # 创建按钮框架并设置背景色
        button_frame = tk.Frame(main_frame, bg='#F0F0F0')
        button_frame.pack()

        # Mac 风格的按钮样式
        button_style = {
            'bg': '#FFFFFF',  # 白色背景
            'fg': '#000000',  # 黑色文字
            'relief': 'solid',  # 实线边框
            'bd': 1,  # 边框宽度
            'highlightthickness': 0  # 去除高亮边框
        }

        # 海运订阅文件
        self.subscription_button = tk.Button(
            button_frame, 
            text="选择海运订阅文件", 
            command=self.select_subscription_file, 
            width=button_width, 
            height=button_height,
            **button_style
        )
        self.subscription_button.grid(row=0, column=0, pady=10, sticky="w")
        self.subscription_ok = tk.Label(
            button_frame, 
            text="未选择文件", 
            fg="red", 
            width=label_width,
            bg='#F0F0F0',  # 设置标签背景色
            font=('Arial', 9)  # 设置字体大小为9
        )
        self.subscription_ok.grid(row=0, column=1, pady=10, padx=(10, 0), sticky="w")

        # 预对账文件
        self.input_button = tk.Button(
            button_frame, 
            text="选择预对账文件", 
            command=self.select_input_file, 
            width=button_width, 
            height=button_height,
            **button_style
        )
        self.input_button.grid(row=1, column=0, pady=10, sticky="w")
        self.input_ok = tk.Label(
            button_frame, 
            text="未选择文件", 
            fg="red", 
            width=label_width,
            bg='#F0F0F0',
            font=('Arial', 9)  # 设置字体大小为9
        )
        self.input_ok.grid(row=1, column=1, pady=10, padx=(10, 0), sticky="w")

        # 输出文件
        self.output_button = tk.Button(
            button_frame, 
            text="选择保存位置", 
            command=self.select_output_file, 
            width=button_width, 
            height=button_height,
            **button_style
        )
        self.output_button.grid(row=2, column=0, pady=10, sticky="w")
        self.output_ok = tk.Label(
            button_frame, 
            text="未选择文件", 
            fg="red", 
            width=label_width,
            bg='#F0F0F0',
            font=('Arial', 9)  # 设置字体大小为9
        )
        self.output_ok.grid(row=2, column=1, pady=10, padx=(10, 0), sticky="w")

        # 开始分析按钮
        self.analyze_button = tk.Button(
            button_frame,
            text="开始分析",
            command=self.start_analysis,
            width=button_width,
            height=button_height,
            **button_style
        )
        self.analyze_button.grid(row=3, column=0, pady=10, sticky="w")

        # 处理标志
        self.processing_done = threading.Event()

        # 创建状态标签
        self.status_label = tk.Label(
            main_frame,
            text="准备就绪",
            fg="#666666",  # 深灰色文字
            bg='#F0F0F0',
            font=('SF Pro Text', 12)  # 使用 macOS 风格的字体
        )
        self.status_label.pack(pady=(10, 0))

    def select_input_file(self):
        self.input_file = filedialog.askopenfilename(
            title="选择预对账文件",
            filetypes=[
                ("Excel 文件", "*.xlsx"),
                ("旧版 Excel 文件", "*.xls"),
                ("所有文件", "*")
            ],
            initialdir="~"
        )
        if self.input_file:
            self.input_ok.config(text="✓ 已选择", fg="green")
        else:
            self.input_ok.config(text="未选择文件", fg="red")

    def select_subscription_file(self):
        self.subscription_file = filedialog.askopenfilename(
            title="选择海运订阅文件",
            filetypes=[
                ("Excel 文件", "*.xlsx"),
                ("旧版 Excel 文件", "*.xls"),
                ("所有文件", "*")
            ],
            initialdir="~"
        )
        if self.subscription_file:
            self.subscription_ok.config(text="✓ 已选择", fg="green")
        else:
            self.subscription_ok.config(text="未选择文件", fg="red")

    def select_output_file(self):
        self.output_file = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel 文件", "*.xlsx"),
                ("所有文件", "*")
            ],
            initialdir="~",
            initialfile="分析结果.xlsx"
        )
        if self.output_file:
            self.output_ok.config(text="✓ 已选择", fg="green")
        else:
            self.output_ok.config(text="未选择文件", fg="red")

    def _update_gui(self, text):
        """
        更新GUI显示的状态文本
        """
        self.status_label.config(text=text)
        self.master.update()

    def _enable_button(self):
        """重新启用按钮"""
        self.analyze_button.config(state='normal')

    def start_analysis(self):
        self.is_running = True
        self.start_time = time.time()  # 记录开始时间
        
        # 禁用开始分析按钮
        self.analyze_button.config(state='disabled')
        
        # 检查必要的文件是否已选择
        if not self.subscription_file:
            self.show_error("错误", "请先选择海运订阅文件")
            self.analyze_button.config(state='normal')
            return
        if not self.output_file:
            self.show_error("错误", "请选择输出文件")
            self.analyze_button.config(state='normal')
            return

        def update_progress(text):
            if not self.is_running:  # 检查是否应该继续运行
                raise Exception("用户取消了操作")
            print(text, flush=True)
            # 使用 after 方法在主线程中更新 GUI
            self.master.after(0, lambda: self._update_gui(text))

        def _process_data():
            try:
                print("开始数据分析...", flush=True)
                
                update_progress("开始读取文件...")
                analyze_data.analyze_excel_data(
                    self.input_file, 
                    self.output_file, 
                    self.subscription_file,
                    status_callback=update_progress
                )
                
                if self.is_running:
                    self.processing_time = time.time() - self.start_time  # 计算处理时间
                    print(f"处理完成，用时 {self.processing_time:.2f} 秒", flush=True)
                    self.master.after(0, lambda: self.show_info("完成", "数据分析已完成！"))
                    update_progress("准备就绪")
                
            except Exception as e:
                if str(e) != "用户取消了操作" and self.is_running:
                    self.master.after(0, lambda: self.show_error("错误", f"处理过程中出现错误：\n{str(e)}"))
                    update_progress("处理出错")
                print(f"操作终止: {str(e)}", flush=True)
            finally:
                self.is_running = False
                self.processing_done.set()
                self.master.after(0, self._enable_button)

        # 创建并启动工作线程
        thread = threading.Thread(target=_process_data)
        thread.daemon = True  # 设置为守护线程，这样主窗口关闭时线程会自动终止
        thread.start()

    def show_error(self, title, message):
        if self.master.winfo_exists():  # 检查窗口是否还存在
            messagebox.showerror(title, message)

    def show_warning(self, title, message):
        if self.master.winfo_exists():  # 检查窗口是否还存在
            messagebox.showwarning(title, message)

    def show_info(self, title, message):
        if self.master.winfo_exists():  # 检查窗口是否还存在
            messagebox.showinfo(title, message)

    def on_closing(self):
        """处理窗口关闭事件"""
        self.is_running = False
        if self.processing_done:
            self.processing_done.set()  # 设置事件
        self.master.destroy()  # 关闭窗口

def run_gui():
    root = tk.Tk()
    gui = DataAnalysisGUI(root)
    
    # 添加一个事件来跟踪处理是否完成
    gui.processing_done = threading.Event()
    
    def on_closing():
        gui.is_running = False
        gui.processing_done.set()  # 设置事件
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
    
    # 等待处理完成
    gui.processing_done.wait()
    
    return gui.input_file, gui.output_file, gui.subscription_file

if __name__ == "__main__":
    input_file, output_file, subscription_file = run_gui()
    print(f"预对账文件：{input_file}")
    print(f"海运订阅文件：{subscription_file}")
    print(f"输出文件：{output_file}")
