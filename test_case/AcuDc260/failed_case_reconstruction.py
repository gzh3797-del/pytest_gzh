import pandas as pd
import os
from pathlib import Path


def process_failed_test_cases(input_files, output_file_path):
    """
    处理262精度测试负电流测试结果文件，筛选失败的行并整理成指定格式

    Parameters:
    input_files (str or list): 输入文件路径或文件路径列表
    output_file_path (str): 输出文件路径
    """
    try:
        # 统一处理输入文件参数
        if isinstance(input_files, str):
            input_files = [input_files]

        all_failed_data = []

        for input_file in input_files:
            print(f"正在处理文件: {input_file}")

            # 检查文件是否存在
            if not os.path.exists(input_file):
                print(f"警告: 文件 {input_file} 不存在，跳过")
                continue

            try:
                # 读取输入文件
                df_input = pd.read_excel(input_file, sheet_name='Sheet')

                # 筛选包含失败的行
                failed_rows = df_input[
                    (df_input['电流2精度测试结果'] == 'Failed') |
                    (df_input['功率1精度测试结果'] == 'Failed') |
                    (df_input['功率2精度测试结果'] == 'Failed') |
                    (df_input['功率总精度测试结果'] == 'Failed')
                    ].copy()

                print(f"  在文件 {os.path.basename(input_file)} 中找到 {len(failed_rows)} 个失败的测试用例")

                # 固定参数
                wait_time = 0  # 等待时间(h)
                sample_count = 20  # 采样次数
                sample_interval = 0.1  # 采样间隔

                # 遍历失败的行
                for _, row in failed_rows.iterrows():
                    test_case = row['测试用例']
                    voltage = row['电压输入值']
                    current_1 = row['电流1输入值']
                    current_2 = row['电流2输入值']

                    # 根据电流值确定精度要求
                    if current_1 in [0.4, -0.4]:
                        voltage_accuracy = 0.001
                        current_accuracy = 0.075
                        power_accuracy = 0.075
                    elif current_1 in [1.5, -1.5]:
                        voltage_accuracy = 0.001
                        current_accuracy = 0.02
                        power_accuracy = 0.02
                    elif current_1 in [3, -3, 5, -5]:
                        voltage_accuracy = 0.001
                        current_accuracy = 0.01
                        power_accuracy = 0.01
                    else:
                        # 默认值
                        voltage_accuracy = 0.001
                        current_accuracy = 0.005
                        power_accuracy = 0.005

                    all_failed_data.append({
                        'test_case': test_case,
                        'Voltage': voltage,
                        'Current_1': current_1,
                        'Current_2': current_2,
                        '等待时间(h)': wait_time,
                        'voltage_accuracy': voltage_accuracy,
                        'current_accuracy': current_accuracy,
                        'power_accuracy': power_accuracy,
                        '采样次数': sample_count,
                        '采样间隔': sample_interval,
                        '来源文件': os.path.basename(input_file)  # 添加来源文件信息
                    })

            except Exception as e:
                print(f"  处理文件 {input_file} 时出错: {str(e)}")
                continue

        if not all_failed_data:
            print("在所有文件中均未找到失败的测试用例")
            return None

        # 创建输出DataFrame
        df_output = pd.DataFrame(all_failed_data)

        # 确保列的顺序正确
        column_order = [
            'test_case', 'Voltage', 'Current_1', 'Current_2', '等待时间(h)',
            'voltage_accuracy', 'current_accuracy', 'power_accuracy',
            '采样次数', '采样间隔', '来源文件'
        ]
        df_output = df_output[column_order]

        # 写入输出文件
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            df_output.to_excel(writer, sheet_name='Failed_Test_Cases', index=False)

        total_failed = len(all_failed_data)
        print(f"\n处理完成！共找到 {total_failed} 个失败的测试用例")
        print(f"结果已保存到: {output_file_path}")

        return df_output

    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        return None


def process_failed_test_cases_from_folder(folder_path, output_file_path, file_pattern="*精度测试*.xlsx"):
    """
    处理文件夹中所有匹配的文件

    Parameters:
    folder_path (str): 文件夹路径
    output_file_path (str): 输出文件路径
    file_pattern (str): 文件匹配模式
    """
    try:
        folder = Path(folder_path)
        if not folder.exists():
            print(f"文件夹 {folder_path} 不存在")
            return None

        # 查找所有匹配的文件
        input_files = list(folder.glob(file_pattern))
        if not input_files:
            print(f"在文件夹 {folder_path} 中未找到匹配 {file_pattern} 的文件")
            return None

        print(f"找到 {len(input_files)} 个匹配的文件:")
        for file in input_files:
            print(f"  - {file.name}")

        # 调用主处理函数
        return process_failed_test_cases([str(file) for file in input_files], output_file_path)

    except Exception as e:
        print(f"处理文件夹时出错: {str(e)}")
        return None


# 使用示例
if __name__ == "__main__":
    # 方法2: 处理多个文件
    print("\n=== 处理多个文件 ===")
    input_files = [
        r"C:\test_data\精度测试结果\262精度测试正电流测试结果.xlsx",
    ]
    result2 = process_failed_test_cases(input_files, r"C:\test_data\精度测试结果\失败测试用例汇总.xlsx")

    # 方法3: 处理整个文件夹
    if result2 is not None:
        print("\n多个文件处理结果预览:")
        print(result2.head())
