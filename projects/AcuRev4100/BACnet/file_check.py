import pandas
from comm.modbus_rtu_tcp import *
from modbus_config import modbus_config

# 设备Modbus地址表中的参数数据类型，后续用于解码
Date_Type_struct_Map = {
    'uint8_t':'!B',
    'uint16_t':'!H',
    'uint32_t':'!I',
    'float32':'!f',
    'double':'!d'
}

def object_change_format(text):
    '''将Object Type转换到与nigara 网站一致'''
    # 去掉 Object 前缀
    if text.startswith("Object "):
        text = text[len("Object "):]

    words = text.split()

    if not words:
        return ""

    # 第一个单词小写，其余首字母大写
    return words[0].lower() + "".join(w.capitalize() for w in words[1:])

def address_config_update(file_path, new_file_path):
    '''更新BACnet 地址表的内容，合并Object Type和ID到一列，删除多余列'''
    address_table = pandas.read_excel(file_path, sheet_name=None)
    result_list = []
    for sheet_name, df in address_table.items():
        try:
            # 防止原数据被修改
            df_copy = df.copy()

            # Object_Type 转换到目标格式后，和ID列合并到Merged_col
            col_a = df["Object Type"].astype(str).apply(object_change_format)
            col_b = df["ID"].astype("Int64").astype(str).replace("<NA>", "")

            merged_col = col_a + ":" + col_b
            df_copy["Object ID"] = merged_col

            result_list.append(df_copy)

        except KeyError as e:
            print(f"缺少列名：{e}，无法处理工作表 {sheet_name}")
        except Exception as e:
            print(f"处理工作表 {sheet_name} 时出错: {e}")

    # 合并所有sheet 的数据，去除不要的列
    final_df = pandas.concat(result_list, ignore_index=True)
    # 写入新的表格，只写一个sheet
    final_df.to_excel(new_file_path, index=False)

# 删除空行
def drop_nan_row(file_path):
    '''删除空行'''
    # 读取 Excel
    df = pandas.read_excel(file_path)
    # 删除 ID 为空的整行
    # 这里处理 NaN、空字符串以及可能的 "nan" 字符串
    df = df[
        df["ID"].notna() & (df["ID"] != "") & (df["ID"] != "nan")
        ]
    # 写回 Excel（覆盖原文件或另存）
    df.to_excel(file_path, index=False)

def drop_nan_col(file_path):
    '''删除空列'''
    df = pandas.read_excel(file_path)
    df = df.dropna(axis=1, how='all')
    df.to_excel(file_path, index=False)

def drop_target_col(file_path, drop_col_name):
    '''删除已知列名的列'''
    df = pandas.read_excel(file_path)
    df = df.drop(columns=drop_col_name)
    df.to_excel(file_path, index=False)

def check_different(row, address_name, point_name, tolerance=0.1, eps=1e-4):
    """
    函数简要描述:处理表格，检查已知列名的两列数据内容，是否一致
    详细描述:
    1. 检查表格中已知两列的数据
    2. 如果是数据值，差值比例不能超过tolerance，如果是极小值和0作比较，差值不能超过eps
    参数：
    param1:表格的每一行
    param2:address_name 检查的列名1
    param3:point_name 检查的列名2
    param4：tolerance 数据类型参数的相差的比例限度值
    param5：eps，极小值或0值之间的比较的差值限度值
    return:比较结果：not found、same、different；并输出到表格
    ExceptionType:数据值错误，或数据类型错误
    """
    val_addr = row[address_name]
    val_point = row[point_name]

    # NaN 判断
    if pandas.isna(val_addr) or pandas.isna(val_point):
        return "not found"

    # 尝试数值比较
    try:
        a = float(val_addr)
        b = float(val_point)

        # 最大误差 = 绝对误差或相对误差
        max_err = max(eps, tolerance * max(abs(a), abs(b)))

        return "same" if abs(a - b) <= max_err else "different"

    except (ValueError, TypeError):
        # 非数值 → 字符串比较
        return "same" if str(val_addr) == str(val_point) else "different"

def check_config_by_object_id(address_file_path, point_file_path, check_address_name, check_point_name, result_file_path):
    """
    函数简要描述:比较两个表格中的目标列内容是否已知
    详细描述:
    1. 两个表格文件中Object ID作为唯一标识，已address 文件为准，合并两个文件，被合并的文件只合入检查列
    2. 检查目标的两列的内容是否相同
    参数：
    param1:address 文件路径
    param2:point 文件路径
    param3:address 检查列名
    param4:point 检查列名
    param5:输出结果的文件路径
    return:将比较的结果输出到文件
    ExceptionType:
    """
    # 读取需要比较的两个表格
    address_df = pandas.read_excel(address_file_path)
    point_df = pandas.read_excel(point_file_path)
    point_df = point_df.dropna(axis=1, how='all')

    merged_df = pandas.merge(
        address_df,
        point_df[["Object ID", check_point_name]],
        on="Object ID",
        how="left",
        suffixes=("_address", "_point")
    )

    if (check_address_name == check_point_name):
        check_address_name = check_address_name + '_address'
        check_point_name = check_point_name + '_point'

    merged_df["compare_result"] = merged_df.apply(
        check_different,
        axis=1,
        args=(check_address_name, check_point_name)
    )

    merged_df.to_excel(result_file_path, index=False)

def analysis_message_to_value(memory_value, data_type):
    '''将报文解析到目标数据类型'''
    try:
        bytes_value = []
        for value in memory_value:
            higt_byte = (value & 0xff00) >> 8
            low_byte = (value & 0xff)
            bytes_value.extend([higt_byte, low_byte])
        if data_type not in Date_Type_struct_Map.keys():
            logging.error('The current data type has not been entered.please enter!')
        value_measu = struct.unpack(Date_Type_struct_Map[data_type], bytes(bytes_value))[0]
        return value_measu
    except ValueError as ve:
        logging.error(f"数据处理错误: {ve}")
    except KeyError as ke:
        logging.error(f"字典查找错误: {ke}")
    except struct.error as se:
        logging.error(f"结构体解析错误: {se}")
    except Exception as e:
        logging.error(f"未知错误: {e}")


def read_value_to_address_table(address_file_path):
    '''根据Address 文件中的Start(Dec)和Reg配置，读取电表中的参数的实时值并写文文档'''
    try:
        # 初始化 Modbus 客户端
        ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    except Exception as e:
        logging.error(f"读取 Excel 文件或初始化 Modbus 客户端时出错: {e}")
        return

    df = pandas.read_excel(address_file_path)
    for index,row in df.iterrows():
        try:
            # 读取电表中的参数值
            value_message = ModbusClient.read_measurement(
                address=int(row['Start(Dec)']),
                count=int(row['Reg']),
                device_id=1
            )
            print(f'转换的数值：{value_message}')

            # 将报文解析为目标数据类型
            value_measu = analysis_message_to_value(value_message, data_type=row['Data type'])
            if value_measu is not None:
                df.at[index, 'value'] = value_measu
            else:
                logging.warning(f"第 {index + 1} 行数据解析失败，跳过该行")

        except KeyError as ke:
            logging.error(f"在第 {index + 1} 行中，缺少必需的列: {ke}")
        except ValueError as ve:
            logging.error(f"第 {index + 1} 行的值转换错误: {ve}")
        except Exception as e:
            logging.error(f"第 {index + 1} 行数据处理时出错: {e}")

    df.to_excel(address_file_path, index=False)