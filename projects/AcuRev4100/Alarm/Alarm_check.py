from comm.source_control import *
import pandas
from modbus_message_switch import *
import datetime

target_alarm_log_datatype_map = {
    'Alarm Log ID':'uint16_t',
    'Alarm Log Time Stamp': {
        'year':'uint8_t',
        'month':'uint8_t',
        'day':'uint8_t',
        'hour':'uint8_t',
        'min':'uint8_t',
        'sec':'uint8_t',
        'millisec':'uint16_t'
    },
    'Alarm Log Monitor ID':'uint16_t',
    'Alarm Status':'uint16_t',
    'Alarm Log Extreme Value 1':'float32',
    'Alarm Log Extreme Value 2':'float32',
    'Alarm Log Duration':'uint32_t'
}

def control_source_by_alarm_channel(alarm_setting_file_path, channel_id, target_status):
    """根据传参以及Alarm Logic check sheet 页的配置，进行控源"""
    alarm_check_df = pandas.read_excel(alarm_setting_file_path, sheet_name='Alarm Logic check')
    for index,row in alarm_check_df.iterrows():
        if int(row['check_Alarm_channel']) == channel_id and row['target_status'] == target_status:
            ua = row['Phase_A_Voltage']
            ub = row['Phase_B_Voltage']
            uc = row['Phase_C_Voltage']
            ia = row['Phase_A_Current']
            ib = row['Phase_B_Current']
            ic = row['Phase_C_Current']
            phase_ua = row['Va_angle']
            phase_ub = row['Vb_angle']
            phase_uc = row['Vc_angle']
            phase_ia = row['Ia_angle']
            phase_ib = row['Ib_angle']
            phase_ic = row['Ic_angle']
            frequency = row['frequency']
            keep_time = int(row['keep_time(s)'])
            print(f"当前Alarm 控源为Ua：{ua}，Ub:{ub},Uc:{uc},Ia:{ia},Ib:{ib},Ic:{ic},Phase_ua:{phase_ua},Phase_ub:{phase_ub},Phase_uc:{phase_uc},Phase_ia:{phase_ia},Phase_ib:{phase_ib},Phase_ic:{phase_ic},frequency:{frequency},Keep_time:{keep_time}s")
            set_ac(
                phase_uc, phase_ub, phase_ua,
                phase_ic, phase_ib, phase_ia,
                uc, ub, ua, ic, ib, ia,frequency
            )
            time.sleep(keep_time)
            break

class Alarm_log():
    def __init__(self,alarm_setting_file_path, alarm_read_sheet_name,client):
        self.alarm_df = pandas.read_excel(alarm_setting_file_path, sheet_name=alarm_read_sheet_name)
        self.client = client
        self.alarm_log_message = 0
        self.alarm_log = 0
        self.single_alarm_log_length = 13

    def get_alarm_log_number(self):
        """获取已经生成的Alarm Log 个数"""
        match_row = self.alarm_df[self.alarm_df['Description'] == 'Number of the Alarm Log']
        number_of_alarm_log_start = int(match_row['Start(Dec)'].values[0])
        reg = int(match_row['Reg'].values[0])
        datatype = match_row['Data type'].values[0]

        number_of_alarm_log = self.client.read_measurement(
            address=number_of_alarm_log_start,
            count=reg,
            device_id=1
        )
        number_of_alarm_log = analysis_message_to_value(number_of_alarm_log, datatype)
        print(f'Alarm Log 当前的数量为{number_of_alarm_log}')
        return number_of_alarm_log

    def set_alarm_log_read_number(self, first_index, read_num):
        """配置读取Alarm日志的配置，从first_index开始，读取read_num条日志"""
        first_index_match_row = self.alarm_df[self.alarm_df['Description'] == 'The first index of the alarm Log will be read']
        first_index_start = int(first_index_match_row['Start(Dec)'].values[0])
        first_index_reg = int(first_index_match_row['Reg'].values[0])
        first_index_datatype = first_index_match_row['Data type'].values[0]
        first_index_bytelist = analysis_value_to_bytelist(first_index, first_index_datatype)
        ret = self.client.write_registers(
            address=first_index_start,
            values=first_index_bytelist,
            device_id=1
        )
        target_ret = f'({first_index_start},{first_index_reg})'
        if target_ret not in str(ret):
            logging.error('set first index fail, ret is:{}'.format(ret))
            return False

        read_num_match_row = self.alarm_df[self.alarm_df['Description'] == 'Read the number of the Alarm Log offset ']
        read_num_start = int(read_num_match_row['Start(Dec)'].values[0])
        read_num_reg = int(read_num_match_row['Reg'].values[0])
        read_num_datatype = read_num_match_row['Data type'].values[0]
        read_num_bytelist = analysis_value_to_bytelist(read_num, read_num_datatype)
        ret = self.client.write_registers(
            address=read_num_start,
            values=read_num_bytelist,
            device_id=1
        )
        target_ret = f'({read_num_start},{read_num_reg})'
        if target_ret not in str(ret):
            logging.error('set read number fail, ret is:{}'.format(ret))
            return False
        return True


    def get_alarm_log_state(self):
        """获取Alarm Log 的状态；0x00: window get data failed;0x0B: window data effective;0xFF: window data Busy;"""
        match_row = self.alarm_df[self.alarm_df['Description'] == 'Data state of Alarm Log Window']
        alarm_log_state_start = int(match_row['Start(Dec)'].values[0])
        reg = int(match_row['Reg'].values[0])
        datatype = match_row['Data type'].values[0]

        alarm_log_state = self.client.read_measurement(
            address=alarm_log_state_start,
            count=reg,
            device_id=1
        )
        alarm_log_state = analysis_message_to_value(alarm_log_state, datatype)
        print(f'Alarm Log 当前的状态为{alarm_log_state}')
        return alarm_log_state

    def analysis_alarm_log_message(self, alarm_log_message, alarm_channel):
        length = self.single_alarm_log_length
        alarm_log_list = [alarm_log_message[i:i+13] for i in range(0, len(alarm_log_message), length)]
        for target_alarm_log_message in alarm_log_list:
            if target_alarm_log_message[5] == alarm_channel:
                logging.info(f'根据alarm_channel 找到的日志内容:{target_alarm_log_message}(未按照日志格式解析)')
                break

        target_alarm_log = {}
        target_alarm_log['Alarm Log ID'] = target_alarm_log_message[0]
        target_alarm_log['Alarm Log Time Stamp'] = {
            'year': split_uint16_to_uint8(int(target_alarm_log_message[1]))[0],
            'month': split_uint16_to_uint8(int(target_alarm_log_message[1]))[1],
            'day': split_uint16_to_uint8(int(target_alarm_log_message[2]))[0],
            'hour': split_uint16_to_uint8(int(target_alarm_log_message[2]))[1],
            'min': split_uint16_to_uint8(int(target_alarm_log_message[3]))[0],
            'sec': split_uint16_to_uint8(int(target_alarm_log_message[3]))[1],
            'millisec': target_alarm_log_message[4]
        }
        target_alarm_log['Alarm Log Monitor ID'] = target_alarm_log_message[5]
        target_alarm_log['Alarm Status'] = target_alarm_log_message[6]
        target_alarm_log['Alarm Log Extreme Value 1'] = target_alarm_log_message[7:9]
        target_alarm_log['Alarm Log Extreme Value 2'] = target_alarm_log_message[9:11]
        target_alarm_log['Alarm Log Duration'] = target_alarm_log_message[11:13]

        for key, value in target_alarm_log.items():
            if key == 'Alarm Log Time Stamp':
                value_time_target = {}
                for key_time, value_time in value.items():
                    if type(value_time) is not list:
                        value_time = [value_time]
                    value_time_target[key_time] = analysis_message_to_value(value_time, target_alarm_log_datatype_map[key][key_time])
                target_alarm_log[key] = value_time_target
            else:
                if type(value) is not list:
                    value = [value]
                target_alarm_log[key] = analysis_message_to_value(value, target_alarm_log_datatype_map[key])
        return target_alarm_log

    def read_alarm_log(self, read_number, read_alarm_cahnnel):
        """获取准备好的Alarm Log"""
        match_row = self.alarm_df[self.alarm_df['Description'] == 'Alarm Log Reading Window']
        alarm_log_start = int(match_row['Start(Dec)'].values[0])
        reg = int(read_number* self.single_alarm_log_length)

        try:
            alarm_log_message = self.client.read_measurement(
                address=alarm_log_start,
                count=reg,
                device_id=1
            )
        except ValueError as e:
            print(f"alarm log数据读取失败: {e}")
        except Exception as e:
            print(f"读取alarm log发生异常: {e}")

        alarm_log_content = self.analysis_alarm_log_message(alarm_log_message, read_alarm_cahnnel)
        return alarm_log_content



    def get_alarm_status(self):
        """获取20个Alarm的状态：1-start，0-end；以列表形式返回"""
        match_row = self.alarm_df[self.alarm_df['Description'] == 'The alarm status of each alarm channel']
        alarm_status_start = int(match_row['Start(Dec)'].values[0])
        alarm_status_reg = int(match_row['Reg'].values[0])
        alarm_status_datatype = match_row['Data type'].values[0]

        alarm_status = self.client.read_measurement(
            address=alarm_status_start,
            count=alarm_status_reg,
            device_id=1
        )
        alarm_status = analysis_message_to_value(alarm_status, alarm_status_datatype)
        state_str = format(alarm_status, '20b')
        state_bit_list = [bit if bit != ' ' else '0' for bit in state_str]
        state_bit_list.reverse()
        print(f'每个Alarm当前的状态为{state_bit_list}')
        return state_bit_list


    def clear_alarm_log(self, clear_enable):
        """清除Alarm Log"""
        match_row = self.alarm_df[self.alarm_df['Description'] == 'Clear alarm Log']
        clear_alarm_log_start = int(match_row['Start(Dec)'].values[0])
        clear_alarm_log_reg = int(match_row['Reg'].values[0])
        clear_alarm_log_datatype = match_row['Data type'].values[0]
        clear_alarm_log_bytelist = analysis_value_to_bytelist(clear_enable, clear_alarm_log_datatype)
        ret = self.client.write_registers(
            address=clear_alarm_log_start,
            values=clear_alarm_log_bytelist,
            device_id=1
        )
        target_ret = f'({clear_alarm_log_start},{clear_alarm_log_reg})'
        if target_ret not in str(ret):
            logging.error('Clear Alarm log fail, ret is:{}'.format(ret))

    def check_alarm_by_alarm_channel(self, actual_status, channel_id, target_status):
        """根据传参检查对应Channel的Alarm的状态，这个接口感觉可以不要，写再更大的功能里"""
        return int(actual_status[channel_id - 1]) == (1 if target_status == 'start' else 0)

    def check_alarm_log(self, target_alarm_config, target_status, current_alarm_log):
        # 获取当前时间
        now = datetime.datetime.now()
        # 获取当前时间并构建字典
        current_time_dict = {
            'year': now.year % 100,
            'month': now.month,
            'day': now.day,
            'hour': now.hour,
            'min': now.minute,
            'sec': now.second,
            'millisec': int(now.microsecond / 1000)
        }

        if target_status == 'start':
            target_status_num = 1
        else:
            target_status_num = 0

        # 检查Channel ID
        result_channel_id = current_alarm_log['Alarm Log Monitor ID'] == target_alarm_config['Alarm Channel']
        logging.info(f'Alarm_channel_id的检查结果{result_channel_id}')
        # 检查告警状态
        result_alarm_status = int(current_alarm_log['Alarm Status']) == target_status_num
        logging.info(f'Alarm_status 的检查结果{result_alarm_status}')
        # 检查Parameter的当前数值是否满足告警要求
        result_parameter = False
        result_parameter2 = False

        current_parameter1 = current_alarm_log['Alarm Log Extreme Value 1']
        target_parameter1_logic = target_alarm_config['Parameter 1 Logic']
        target_parameter1_pickup_value = target_alarm_config['Parameter 1 Pickup Value']
        expression_parameter1 = f'{current_parameter1}{target_parameter1_logic}{target_parameter1_pickup_value}'
        result_parameter1 = eval(expression_parameter1)
        logging.info(f'Alarm_parameter1 的检查结果{result_parameter1}')

        if target_alarm_config['Channel Logic'] != 'Disable':
            current_parameter2 = current_alarm_log['Alarm Log Extreme Value 2']
            target_parameter2_logic = target_alarm_config['Parameter 2 Logic']
            target_parameter2_pickup_value = target_alarm_config['Parameter 2 Pickup Value']
            expression_parameter2 = f'{current_parameter2}{target_parameter2_logic}{target_parameter2_pickup_value}'
            result_parameter2 = eval(expression_parameter2)
            logging.info(f'Alarm_parameter2 的检查结果{result_parameter2}')

        if target_alarm_config['Channel Logic'] == '&&':
            result_parameter = result_parameter1 and result_parameter2
            logging.info(f'Alarm_parameter 的检查结果{result_parameter}')
        elif target_alarm_config['Channel Logic'] == '||':
            result_parameter = result_parameter1 or result_parameter2
            logging.info(f'Alarm_parameter 的检查结果{result_parameter}')

        # 检查当前告警的触发时间，检查到分钟
        time_year_result = current_time_dict['year'] == current_alarm_log['Alarm Log Time Stamp']['year']
        time_month_result = current_time_dict['month'] == current_alarm_log['Alarm Log Time Stamp']['month']
        time_day_result = current_time_dict['day'] == current_alarm_log['Alarm Log Time Stamp']['day']
        time_hour_result = current_time_dict['hour'] == current_alarm_log['Alarm Log Time Stamp']['hour']
        time_min_result = abs(current_time_dict['min'] == current_alarm_log['Alarm Log Time Stamp']['min']) < 2

        time_result = time_year_result and time_month_result and time_day_result and time_hour_result and time_min_result
        logging.info(f'Alarm_time 的检查结果{time_result}')

        if target_status == 'start':
            result_duration = int(current_alarm_log['Alarm Log Duration']) == 0
        else:
            result_duration = int(current_alarm_log['Alarm Log Duration']) != 0
        logging.info('Alarm_duration 的检查结果为{}'.format(result_duration))

        if result_channel_id and result_alarm_status and result_parameter and time_result and result_duration:
            logging.info('当前告警触发，且所有条件满足预期')
            return True

        logging.info('告警条件未全部预期')
        return False

