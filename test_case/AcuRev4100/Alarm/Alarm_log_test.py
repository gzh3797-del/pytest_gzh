from Alarm_set import *
from Alarm_check import *
from comm.modbus_rtu_tcp import ModbusRtuOrTcp
import pandas
from modbus_message_switch import *

Log(str(__file__).split("\\")[-1])

ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])

# 根据Alarm Setting 表格中，Alarm setting sheet页里每个Alarm的parameter1和parameter2的配置，找到Modbus地址表中的Start(Dec)，再写入这个文件中
parameter_setting_file_path = 'Parameters setting.xlsx'

# Alarm setting的路径
# 1. Alarm setting sheet页准备了20个Alarm的配置
# 2. Alarm Logic check sheet页准备了20个ALarm start和end需要的源输入配置
# 3. Alarm Setting Modbus Address sheet页准备了Alarm 配置下发的地址
# 4. Alarm Reading Modbus Address sheet页准备了AlarmLog 读取相关的地址
alarm_setting_file_path = 'Alarm setting.xlsx'
alarm_setting_sheet_name = 'Alarm setting'
alarm_check_sheet_name = 'Alarm Logic check'
alarm_setting_modbus_sheet_name = 'Alarm Setting Modbus Address'
alarm_read_sheet_name = 'Alarm Reading Modbus Address'
result_file_path = 'result.xlsx'

# 一个channel 的Alarm 配置的寄存器个数
alarm_config_length = 17

# 电表的Modbus地址表，用来查找Alarm 参数的Start(Dec)
modbus_address_file_path = 'AcuRev4100 Modbus Address Table.xlsx'

# Alarm Setting 的modbus 地址与填写内容的对应关系
alarmsetting_Map = {
    'Alarm channel 1 parameter1 index':'Parameter 1 Start(Dec)',
    'Alarm channel 1 parameter2 index':'Parameter 2 Start(Dec)',
    'Alarm channel 1 parameter1 Pickup Value':'Parameter 1 Pickup Value',
    'Alarm channel 1 parameter1 Pickup Delay':'Parameter 1 Pickup Delay',
    'Alarm channel 1 parameter1 Dropout Value':'Parameter 1 Dropout Value',
    'Alarm channel 1 parameter1 Dropout Delay':'Parameter 1 Dropout Delay',
    'Alarm channel 1 parameter2 Pickup Value':'Parameter 2 Pickup Value',
    'Alarm channel 1 parameter2 Pickup Delay':'Parameter 2 Pickup Delay',
    'Alarm channel 1 parameter2 Dropout Value':'Parameter 2 Dropout Value',
    'Alarm channel 1 parameter2 Dropout Delay':'Parameter 2 Dropout Delay',
    'Alarm channel 1 DO  selection':'DO selection'
}

if __name__ == '__main__':
    logging.info('根据Alarm setting.xlsx 文件中的Alarm setting sheet页的配置，查找Parameter index在Modbus地址表中对应的Start地址，后续下发参数时使用')
    get_parameter_index_start(
        alarm_setting_file_path,
        alarm_setting_sheet_name,
        modbus_address_file_path,
        parameter_setting_file_path)

    logging.info('创建Alarm_setting 类，用来配置ALarm参数')
    alarm_setting = Alarm_parameter(
        parameter_setting_file_path,
        alarm_setting_file_path,
        alarm_setting_modbus_sheet_name,
        Alarm_config_length,
        ModbusClient
    )

    logging.info('根据Parameter_setting_output.xlsx 文件中的配置，下发Enable和Logic以外的所有参数')
    for key, value in alarmsetting_Map.items():
        logging.info('配置Alarm的所有{}'.format(value))
        alarm_setting.set_Alarm_parameter(
            channel1_parameter_name=key,
            value_name=value
        )
    logging.info('配置20个Alarm的Logic参数')
    alarm_setting.set_Alarm_Logic()
    logging.info('配置20个Alarm的Enable参数')
    alarm_setting.set_Alarm_Enable()

    '''
    1. 遍历Alarm setting.xlsx 的Alarm Logic check sheet页面
    2. 清除Alarm Log 
    3. 根据配置控源并等待足够的时间
    4. 获取alarm state，并检查对应channel 的状态是否符合预期
    5. 获取Alarm Log 的数量
    6. 配置计划读取的Alarm log 的first index，以及计划读取的Alarm Log 数量
    7. 持续获取Alarm Log 的准备状态，如果等于0xB，也就是11，则说明准备好了日志，可以读取了
    8. 根据num 和index 配置读取已经准备好的Alarm log
    '''
    # 尝试控源，切换至自动模式，下发指令，将源控制输出为全0
    switch_device_screen_interface(0x01)
    time.sleep(5)
    set_ac(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    time.sleep(5)

    alarm_check_df = pandas.read_excel(alarm_setting_file_path, sheet_name=alarm_check_sheet_name)
    # 转换 'check_result' 列的类型为 object
    alarm_check_df['check_result'] = alarm_check_df['check_result'].astype('object')

    alarm_log = Alarm_log(alarm_setting_file_path, alarm_read_sheet_name, ModbusClient)

    for index, row in alarm_check_df.iterrows():
        # 每次构造Alarm触发或结束的场景之前，先清除Alarmlog。确保Alarm Log 的数量不会太多
        alarm_log.clear_alarm_log(clear_enable=1)

        alarm_channel_id = row['check_Alarm_channel']
        target_alarm_status = row['target_status']
        print('-------------------------当前触发的Alarm 的Channel ID 为:{}-------------------------'.format(alarm_channel_id))
        print('-------------------------当前触发的Alarm 的target status 为:{}----------------------'.format(target_alarm_status))
        print('根据当前的alarm_channel_id获取Alarm setting 配置:')
        alarm_setting_config = alarm_setting.get_alarm_setting_by_alarm_channel_id(alarm_channel_id)
        print(alarm_setting_config)

        control_source_by_alarm_channel(alarm_setting_file_path, alarm_channel_id, target_alarm_status)

        status = alarm_log.get_alarm_status()
        check_state = alarm_log.check_alarm_by_alarm_channel(status, alarm_channel_id, target_alarm_status)
        print('channel {} Alarm 当前的状态检查情况为{}'.format(alarm_channel_id, check_state))

        # 获取Alarm Log 的数量
        alarm_log_number = alarm_log.get_alarm_log_number()

        # 配置计划读取的第1个Alarm Log 的index默认为1，以及计划读取的Alarm Log 数量为当前全部长度
        alarm_first_index = 1
        # print("配置计划读取的第1个Alarm Log 的index默认为1，如果当前Alarm log数量超过10，index配置Alarmlog数量的一半，以及计划读取的Alarm Log 数量为当前全部长度")
        # if alarm_log_number > 10:
        #     alarm_first_inde = alarm_log_number // 2
        print('配置alarm_first_index为{}，alarm_log_number为{}'.format(alarm_first_index,alarm_log_number ))
        alarm_log.set_alarm_log_read_number(first_index=alarm_first_index, read_num=alarm_log_number)

        print("持续获取Alarm log 的准备状态，如果等于0x0B(11)，则说明准备好了日志，可以读取了")
        alarm_log_state = 0
        for i in range(10):
            time.sleep(1)
            alarm_log_state = alarm_log.get_alarm_log_state()
            if alarm_log_state == 11:
                break

        if alarm_log_state != 11:
            print('当前Alarm未生成日志，无法进行后续步骤')
            alarm_check_df.at[index, 'check_result'] = 'FAILED'
            continue

        print("读取根据num 和 channel id 配置已经准备好的Alarm log：")
        alarm_log_content = alarm_log.read_alarm_log(read_number=alarm_log_number,
                                                     read_alarm_cahnnel=alarm_channel_id)
        print(alarm_log_content)
        check_result = alarm_log.check_alarm_log(alarm_setting_config, target_alarm_status, alarm_log_content)
        if check_result:
            alarm_check_df.at[index, 'check_result'] = 'PASS'
        else:
            alarm_check_df.at[index, 'check_result'] = 'FAILED'


    alarm_check_df.to_excel(result_file_path, index=False)

    set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
    time.sleep(5)
    switch_device_screen_interface(0x00)



