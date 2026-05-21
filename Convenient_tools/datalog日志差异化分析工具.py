import pandas as pd
import numpy as np
import re
import tkinter as tk
from tkinter import filedialog, ttk, messagebox


def extract_numeric_value(value):
    """从可能带有单位的字符串中提取数值部分"""
    # 如果已经是数值类型，直接返回
    if pd.api.types.is_numeric_dtype(type(value)):
        return float(value), True

    # 转换为字符串处理
    str_val = str(value).strip()

    # 正则表达式匹配数字（包括整数、小数、负数）
    match = re.search(r'^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?', str_val)

    if match:
        # 提取并转换为浮点数
        num_str = match.group()
        try:
            return float(num_str), True
        except ValueError:
            return value, False
    else:
        # 没有找到数字
        return value, False


def check_row_differences(csv_file, threshold=0.2, abs_threshold=None):
    """检查CSV文件中每行数据与上一行的差异（从第三列开始），返回差异较大的行"""
    try:
        # 读取CSV文件
        df = pd.read_csv(csv_file)

        if df.shape[0] < 2:
            return None, "文件数据不足，至少需要两行数据进行比较"

        large_diff_rows = []

        # 遍历每一行（从第二行开始，与上一行比较）
        for idx, row in df.iterrows():
            if idx == 0:  # 跳过第一行，从第二行开始比较
                continue

            # 获取上一行数据作为基准
            prev_row = df.iloc[idx - 1]
            row_diff_info = {
                '行索引': idx + 1,
                '上一行索引': idx,
                '差异列数': 0,
                '差异详情': []
            }

            # 比较每一列（从第三列开始，索引为2）
            for col_idx, col in enumerate(df.columns):
                # 从第三列开始比较（索引2及以后）
                if col_idx < 2:
                    continue

                # 获取当前列的上一行值和当前值
                prev_val = prev_row[col]
                current_val = row[col]

                # 处理NaN值
                if pd.isna(prev_val) or pd.isna(current_val):
                    if not (pd.isna(prev_val) and pd.isna(current_val)):
                        row_diff_info['差异列数'] += 1
                        row_diff_info['差异详情'].append(
                            f"{col} (列{col_idx + 1}): 上一行值为{prev_val}，当前值为{current_val}（存在NaN值）"
                        )
                    continue

                # 提取数值部分
                prev_num, prev_is_numeric = extract_numeric_value(prev_val)
                current_num, current_is_numeric = extract_numeric_value(current_val)

                # 判断是否都是可提取数值的类型
                if prev_is_numeric and current_is_numeric:
                    # 数值类型：计算差异
                    abs_diff = abs(current_num - prev_num)
                    is_large_diff = False
                    rel_diff = 0

                    # 检查相对差异
                    if prev_num != 0:
                        rel_diff = abs_diff / abs(prev_num)
                        if rel_diff > threshold:
                            is_large_diff = True
                    elif abs_diff > 0:
                        is_large_diff = True

                    # 检查绝对差异（如果设置了）
                    if abs_threshold is not None and abs_diff > abs_threshold:
                        is_large_diff = True

                    if is_large_diff:
                        row_diff_info['差异列数'] += 1
                        row_diff_info['差异详情'].append(
                            f"{col} (列{col_idx + 1}): 上一行='{prev_val}'({prev_num}), 当前='{current_val}'({current_num}), "
                            f"绝对差={abs_diff:.4f}, 相对差={rel_diff:.2%}"
                        )
                else:
                    # 非数值类型：直接判断是否相等（比较原始值）
                    if str(prev_val) != str(current_val):
                        row_diff_info['差异列数'] += 1
                        row_diff_info['差异详情'].append(
                            f"{col} (列{col_idx + 1}): 上一行值='{prev_val}', 当前值='{current_val}'（不相等）"
                        )

            if row_diff_info['差异列数'] > 0:
                large_diff_rows.append(row_diff_info)

        return large_diff_rows, f"共分析 {df.shape[0]} 行，从第3列开始比较，共涉及 {df.shape[1] - 2} 列，每行与上一行进行比较"

    except Exception as e:
        return None, f"分析错误: {str(e)}"


def select_file():
    """打开文件选择对话框"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择CSV文件",
        filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
    )
    return file_path


def show_results(results, summary):
    """显示分析结果"""
    window = tk.Toplevel()
    window.title("差异分析结果")
    window.geometry("800x600")

    # 显示摘要信息
    ttk.Label(window, text=summary, font=("Arial", 10, "bold")).pack(pady=10)

    # 创建滚动条
    frame = ttk.Frame(window)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    scrollbar = ttk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # 创建文本区域显示结果
    text_widget = tk.Text(frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar.config(command=text_widget.yview)

    # 填充结果
    if not results:
        text_widget.insert(tk.END, "所有行与上一行的差异均在合理范围内")
    else:
        text_widget.insert(tk.END, f"发现 {len(results)} 行与上一行差异较大：\n\n")
        for row in results:
            text_widget.insert(tk.END, f"第{row['行索引']}行（与第{row['上一行索引']}行比较）：\n")
            for detail in row['差异详情']:
                text_widget.insert(tk.END, f"  - {detail}\n")
            text_widget.insert(tk.END, "\n")

    text_widget.config(state=tk.DISABLED)  # 设置为只读


def main():
    # 创建参数设置窗口
    param_window = tk.Tk()
    param_window.title("设置参数")
    param_window.geometry("400x200")

    # 相对阈值设置
    ttk.Label(param_window, text="相对差异阈值 (例如0.2表示20%):").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
    threshold_var = tk.DoubleVar(value=0.2)
    ttk.Entry(param_window, textvariable=threshold_var).grid(row=0, column=1, padx=10, pady=10)

    # 绝对阈值设置
    ttk.Label(param_window, text="绝对差异阈值 (可选，留空则不使用):").grid(row=1, column=0, padx=10, pady=10,
                                                                           sticky=tk.W)
    abs_threshold_var = tk.StringVar(value="")
    ttk.Entry(param_window, textvariable=abs_threshold_var).grid(row=1, column=1, padx=10, pady=10)

    def run_analysis():
        """运行分析"""
        # 获取阈值参数
        try:
            threshold = threshold_var.get()
            abs_threshold = float(abs_threshold_var.get()) if abs_threshold_var.get() else None
        except ValueError:
            messagebox.showerror("参数错误", "绝对阈值必须是数字")
            return

        # 选择文件
        file_path = select_file()
        if not file_path:
            return

        # 执行分析
        results, summary = check_row_differences(file_path, threshold, abs_threshold)

        # 显示结果
        if results is None:
            messagebox.showerror("错误", summary)
        else:
            param_window.destroy()
            show_results(results, summary)

    # 运行按钮
    ttk.Button(param_window, text="选择文件并分析", command=run_analysis).grid(row=2, column=0, columnspan=2, pady=20)

    param_window.mainloop()


if __name__ == "__main__":
    main()
