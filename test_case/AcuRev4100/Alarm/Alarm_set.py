import pandas
from modbus_message_switch import *

Alarm_config_length = 17

parameter_logic_map = {
    'Disable':0,
    'equal':1,
    '>':2,
    '<':3
}
channel_logic_map = {
    'Disable':0,
    '&&':1,
    '||':2
}
enable_map = {
    'Enable':1,
    'Disable':0
}

def get_parameter_index_start(Alarm_setting_file_path, Alarm_setting_sheet_name, Modbus_Address_file_path, parameter_setting_file_path):
    """
    函数简要描述:获取Alarm 参数的Start(Dec)，生成新的Parameter_setting 文件
    参数：
    Alarm_modbus_setting_file_path:Alarm Setting文件存储地址
    Alarm_setting_sheet_name：Alarm Setting 文件中配置告警参数的shhet页名字
    Modbus_Address_file_path：电表Modbus 地址表的地址
    parameter_setting_file_path： 完成处理后，将结果输出到这个文件地址
    return:parameter_setting_file_path
    """
    alarm_setting = pandas.read_excel(Alarm_setting_file_path, sheet_name=Alarm_setting_sheet_name)
    parameter_config = pandas.read_excel(Modbus_Address_file_path, sheet_name=None)

    Parameter_1_index = alarm_setting['Parameter 1 index']
    Parameter_2_index = alarm_setting['Parameter 2 index']

    index_map = {
        'Parameter_1_index':Parameter_1_index,
        'Parameter_2_index':Parameter_2_index
    }
    target_name_map = {
        'Parameter_1_index': 'Parameter 1 Start(Dec)',
        'Parameter_2_index': 'Parameter 2 Start(Dec)'
    }

    for key,value in index_map.items():
        matching_rows = []
        for parameter_value in value:
            found = False
            for sheet_name, parameter_df in parameter_config.items():
                matching_row = parameter_df[parameter_df['Description'] == parameter_value]
                if not matching_row.empty:
                    target_value = matching_row['Start(Dec)'].values[0]
                    matching_rows.append(target_value)
                    found = True
                    break
            if not found:
                matching_rows.append(None)
                logging.error("{} 的Start(Dec) 未找到".format(parameter_value))
        alarm_setting[target_name_map[key]] = matching_rows
        with pandas.ExcelWriter(parameter_setting_file_path, engine='xlsxwriter') as writer:
            alarm_setting.to_excel(writer, sheet_name=Alarm_setting_sheet_name, index=False)


class Alarm_parameter():
    def __init__(self,parameter_setting_file_path,Alarm_modbus_setting_file_path,alarm_setting_modbus_sheet_name,Alarm_config_length, client):
        self.parameter_df = pandas.read_excel(parameter_setting_file_path)
        self.alarm_setting_df = pandas.read_excel(Alarm_modbus_setting_file_path, sheet_name=alarm_setting_modbus_sheet_name)
        self.alarm_config_length = Alarm_config_length
        self.client = client

    def set_Alarm_parameter(self, channel1_parameter_name, value_name):
        """
        函数简要描述:根据传参，配置Alarm 除了Logic和Enable以外的参数
        参数：
        channel1_parameter_name:channel1 Alarm 参数配置的描述
        value_name：预期配置的20个ALarm的参数
        return:配置失败会中断，没有返回
        """
        channel_1_matching_row = self.alarm_setting_df[self.alarm_setting_df['Description'] == channel1_parameter_name]
        channel_1_start = int(channel_1_matching_row['Start(Dec)'].values[0])
        datatype = channel_1_matching_row['Data type'].values[0]
        reg = int(channel_1_matching_row['Reg'].values[0])

        for index, row in self.parameter_df.iterrows():
            alarm_channel = int(row['Alarm Channel'])
            start = channel_1_start + (alarm_channel-1)*self.alarm_config_length
            target_value = row[value_name]
            value = analysis_value_to_bytelist(target_value, datatype)
            ret = self.client.write_registers(address=start, values=value, slave=1)
            ret_right = f"({start},{reg})"
            if ret_right not in str(ret):
                logging.error('set channel {} {} fail, ret is:{}'.format(alarm_channel, value_name, ret))
                return False
        return True

    def set_Alarm_Logic(self):
        """函数简要描述:配置Alarm Logic 参数"""
        channel_1_logic_matching_row = self.alarm_setting_df[self.alarm_setting_df['Description'] == 'Alarm Channel 1 Logic']
        channel_1_logic_start = int(channel_1_logic_matching_row['Start(Dec)'].values[0])
        reg = int(channel_1_logic_matching_row['Reg'].values[0])

        for index, row in self.parameter_df.iterrows():
            alarm_channel = int(row['Alarm Channel'])
            start = channel_1_logic_start + (alarm_channel-1)*self.alarm_config_length
            parameter1_logic = int(parameter_logic_map[row['Parameter 1 Logic']])
            parameter2_logic = int(parameter_logic_map[row['Parameter 2 Logic']])
            channel_logic = int(channel_logic_map[row['Channel Logic']])

            tartget_value = (channel_logic << 6) + (parameter2_logic << 3) + parameter1_logic

            ret = self.client.write_registers(address=start, values=tartget_value, slave=1)
            ret_right = f'({start},{reg})'
            if ret_right not in str(ret):
                logging.error('set channel {} logic fail, ret is:{}'.format(alarm_channel, ret))
                return False
        return True


    def set_Alarm_Enable(self):
        """配置20个Alarm的Enable 参数"""
        enable_matching_row = self.alarm_setting_df[self.alarm_setting_df['Description'] == 'Enable all alarm ']
        start = int(enable_matching_row['Start(Dec)'].values[0])
        reg = int(enable_matching_row['Reg'].values[0])
        datatype = enable_matching_row['Data type'].values[0]

        enable_value_list = []
        for index, row in self.parameter_df.iterrows():
            enable_value = int(enable_map[row['Alarm Enable']])
            enable_value_list.insert(0, enable_value)

        target_value = 0
        for enable in enable_value_list:
            target_value = (target_value << 1) | enable
        print(f'Enable all Alarm parameter is {target_value:20b}')

        value = analysis_value_to_bytelist(target_value, datatype)

        ret = self.client.write_registers(address=start, values=value, slave=1)
        ret_right = f'({start},{reg})'
        if ret_right not in str(ret):
            logging.error('set enable fail, ret is:{}'.format(ret))
            return False
        return True

    def get_alarm_setting_by_alarm_channel_id(self, alarm_channel_id):
        """根据channel_id查找对应Alarm 的配置并返回Alarm的完整配置"""
        alarm_matching_row = self.parameter_df[self.parameter_df['Alarm Channel'] == alarm_channel_id]
        alarm_matching_row_dict = alarm_matching_row.iloc[0].to_dict()
        return alarm_matching_row_dict
