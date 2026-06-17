from nigara_web_control import NigaraWebControl
from file_check import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

# nigara 控制器的url
nigara_url = 'http://192.168.1.140/login'
# nigara控制器服务网页的登录账号与密码
username = 'admin'
password = 'Admin12345'
# BACnet 地址表的路径，如果要检查Value，要在文件中添加参数的Modbus地址表配置：Data Type、Reg、Start(Dec)、End(Dec)
address_table_path = 'AcuRev4100 Bacnet Address Table.xlsx'
# 完成处理的BACnet 地址表的文件路径
address_after_update_path = 'address_new.xlsx'
# 从nigara 网页上获取到的设备参数信息，包含配置与value
point_file_path = 'points_output.xlsx'
# 将完成匹配检查的结果写入到下方路径
result_file_path = 'check_result.xlsx'
# 预期对比检查的地址表中的列名与nigara网页中获取到的信息的列名
address_check_name = "value"
point_check_name = "Value"

if __name__ == '__main__':
    nigara = NigaraWebControl(nigara_url,username,password)
    print('登录nigara 控制器的网页服务')
    nigara.login()
    print('在nigara 网页服务中，查找设备并添加到database，目前只支持查找1个设备')
    nigara.discover_device()
    nigara.add_device()
    print('在nigara 网页服务中，在已经添加成功设备中，查找所有参数信息并输出到excel文件')
    nigara.discover_points()
    nigara.points_info_output(point_file_path)

    print('处理BACnet地址表的格式，将Object Type的格式更新至与nigara服务一致，并与ID列合并')
    address_config_update(address_table_path, address_after_update_path)
    print('删除不需要的列')
    drop_col_name = ["Real Time", "Energy"]
    drop_target_col(address_after_update_path, drop_col_name)
    print('删除BACnet地址表中的空行')
    drop_nan_row(address_after_update_path)
    print('删除 point 文件中的空列')
    drop_nan_col(point_file_path)
    print('根据表格中的Modbus地址表的配置信息，获取电表的实际value')
    read_value_to_address_table(address_after_update_path)
    print('根据object id检查address和point文件中的目标列')
    check_config_by_object_id(address_after_update_path, point_file_path, address_check_name, point_check_name, result_file_path)