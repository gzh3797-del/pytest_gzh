import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext
import os
import re
import threading


class NonZeroFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("非零值查找工具")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)

        # 设置样式
        self.style = ttk.Style()
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("Header.TLabel", font=("SimHei", 12, "bold"))

        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择区域
        self.file_frame = ttk.LabelFrame(self.main_frame, text="文件选择", padding="10")
        self.file_frame.pack(fill=tk.X, pady=(0, 10))

        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path_var, width=60)
        self.file_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)

        self.browse_btn = ttk.Button(self.file_frame, text="浏览...", command=self.browse_file)
        self.browse_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.start_col_frame = ttk.Frame(self.file_frame)
        self.start_col_frame.pack(side=tk.LEFT)

        ttk.Label(self.start_col_frame, text="起始列:").pack(side=tk.LEFT, padx=(0, 5))
        self.start_col_var = tk.StringVar(value="3")
        self.start_col_entry = ttk.Entry(self.start_col_frame, textvariable=self.start_col_var, width=5)
        self.start_col_entry.pack(side=tk.LEFT)

        self.find_btn = ttk.Button(self.file_frame, text="查找非零值", command=self.start_finding)
        self.find_btn.pack(side=tk.LEFT)

        # 结果显示区域
        self.result_frame = ttk.LabelFrame(self.main_frame, text="查找结果", padding="10")
        self.result_frame.pack(fill=tk.BOTH, expand=True)

        self.result_text = scrolledtext.ScrolledText(self.result_frame, wrap=tk.WORD, font=("SimHei", 10))
        self.result_text.pack(fill=tk.BOTH, expand=True)
        self.result_text.config(state=tk.DISABLED)

        # 状态区域
        self.status_var = tk.StringVar(value="就绪")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.main_frame, variable=self.progress_var, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(10, 0))
        self.progress_bar.pack_forget()  # 初始隐藏

    def browse_file(self):
        """打开文件选择对话框"""
        file_path = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[
                ("数据文件", "*.xlsx *.xls *.csv"),
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.file_path_var.set(file_path)

    def log(self, message):
        """在结果区域显示消息"""
        self.result_text.config(state=tk.NORMAL)
        self.result_text.insert(tk.END, message + "\n")
        self.result_text.see(tk.END)  # 滚动到最后
        self.result_text.config(state=tk.DISABLED)

    def set_status(self, status):
        """更新状态条"""
        self.status_var.set(status)

    def extract_numeric_value(self, value):
        """从可能包含单位的字符串中提取数值"""
        if pd.isna(value):
            return None

        # 如果是数字类型直接返回
        if isinstance(value, (int, float)):
            return value

        # 转换为字符串处理
        str_value = str(value).strip()

        # 使用正则表达式提取数字（支持正负号、整数、小数、科学计数法）
        match = re.search(r'[-+]?\d*\.?\d+([eE][-+]?\d+)?', str_value)

        if match:
            try:
                return float(match.group())
            except ValueError:
                return None
        return None

    def find_nonzero_values(self):
        """查找非零值的核心函数"""
        file_path = self.file_path_var.get()
        if not file_path:
            self.log("请先选择文件")
            self.set_status("就绪")
            self.progress_bar.pack_forget()
            return

        try:
            start_col = int(self.start_col_var.get()) - 1  # 转换为0-based索引
            if start_col < 0:
                raise ValueError("起始列不能小于1")
        except ValueError as e:
            self.log(f"输入错误: {str(e)}")
            self.set_status("就绪")
            self.progress_bar.pack_forget()
            return

        file_ext = os.path.splitext(file_path)[1].lower()

        try:
            self.log(f"正在读取文件: {os.path.basename(file_path)}")
            self.set_status("正在读取文件...")

            # 读取文件
            if file_ext in ['.xlsx', '.xls']:
                df = pd.read_excel(file_path)
            elif file_ext == '.csv':
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            else:
                self.log(f"不支持的文件格式: {file_ext}")
                self.set_status("就绪")
                self.progress_bar.pack_forget()
                return

            # 检查起始列是否超出范围
            if start_col >= len(df.columns):
                self.log(f"错误：起始列索引{start_col + 1}超出文件实际列数{len(df.columns)}")
                self.set_status("就绪")
                self.progress_bar.pack_forget()
                return

            # 从指定列开始截取数据
            df = df.iloc[:, start_col:]
            total_rows = len(df)
            total_nonzero = 0

            self.log(f"从第{start_col + 1}列开始的非零值信息（忽略单位）：")
            self.log("-" * 80)
            self.progress_bar.pack(fill=tk.X, pady=(10, 0))
            self.progress_var.set(0)

            # 遍历查找非零值
            for col_idx, col in enumerate(df.columns):
                col_nonzero = 0
                # 每处理一列更新一次进度
                progress = ((col_idx + 1) / len(df.columns)) * 100
                self.progress_var.set(progress)
                self.root.update_idletasks()

                for index, cell_value in df[col].items():
                    # 提取数值部分
                    numeric_value = self.extract_numeric_value(cell_value)

                    # 判断是否为非零数值
                    if numeric_value is not None and numeric_value != 0:
                        col_nonzero += 1
                        total_nonzero += 1
                        self.log(f"行 {index + 1}: {col}, 原始值='{cell_value}', 提取的数值={numeric_value}")

            self.log("-" * 80)
            self.log(f"总计非零值数量: {total_nonzero}")
            self.set_status("查找完成")
            self.progress_var.set(100)

        except Exception as e:
            self.log(f"处理文件时出错: {str(e)}")
            self.set_status("出错")
        finally:
            self.progress_bar.pack_forget()

    def start_finding(self):
        """启动查找线程，避免界面卡顿"""
        # 清空之前的结果
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.config(state=tk.DISABLED)

        # 在新线程中执行查找，防止界面冻结
        threading.Thread(target=self.find_nonzero_values, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = NonZeroFinderApp(root)
    root.mainloop()
