import cmath
import math
import statistics
import struct

from comm.modbus_rtu_tcp import ModbusRtuOrTcp
# from comm.source_control import SourCon
from test_case.AcuRev1320.TOU.tou_memory_addrs import MemoryAddr, SlaveId, MemoryReg
from tools.log import Log
from modbus_config import modbus_config


class HandleMemory:
    def __init__(self, slave_id=SlaveId.slave_id):
        """
        初始化实例
        :param slave_id: 电表标志slave_id
        """
        self.slave_id = slave_id
        self.modbus_client = None
        self.log = None
        self.init_func()

    def init_func(self):
        """
        ModBus连接,log初始化
        :return:
        """
        self.modbus_client = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
        self.log = Log(str(__file__).split("\\")[-1]).logger

    @staticmethod
    def get_bytes_value(memory_value):
        """
        解析寄存器返回值
        :param memory_value: 寄存器返回值
        :return: 整数列表
        """
        bytes_value = []
        for value in memory_value:
            high_byte = (value & 0xff00) >> 8
            low_byte = (value & 0x00ff)
            bytes_value.extend([high_byte, low_byte])
        return bytes_value

    def compare_res_by_set_voltage_wire_mode(self, exp_val, act_val):
        """
        判断电压接线方式是否成功设置
        :param exp_val: 期望值
        :param act_val: 寄存器值
        :return: 判断结果
        """
        if exp_val == act_val:
            self.log.info(f'set_wire_mode_by_voltage pass, act_val is:{act_val}')
            return True
        else:
            self.log.info(f'set_wire_mode_by_voltage fail, exp_val is:{exp_val}, act_val is:{act_val}')
            return False

    def read_enable_tou(self):
        """

        :return: 电表enable_tou值
        """
        address = MemoryAddr.enable_tou
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)

        return measure_value[0]

    def read_dst_enable(self):
        """

        :return: 电表enable_tou值
        """
        address = MemoryAddr.dst_enable
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)

        return measure_value[0]

    def read_enable_special_weekday_schedule(self):
        """

        :return: 电表enable_tou值
        """
        address = MemoryAddr.enable_special_weekday_schedule
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)

        return measure_value[0]

    def read_holiday_setting_enable(self):
        """

        :return: 电表holiday_setting_enable值
        """
        address = MemoryAddr.holiday_setting_enable
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)

        return measure_value[0]

    def read_monthly_billing_mode(self):
        """

        :return: 电表monthly_billing_mode值
        """
        address = MemoryAddr.monthly_billing_mode
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)

        return measure_value[0]

    def read_billing_time(self):
        """

        :return: 电表billing_time值
        """
        address = MemoryAddr.billing_time
        count = 4
        slave = self.slave_id
        lst = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        formatted = f"{lst[0]:02d} {lst[1]:02d}:{lst[2]:02d}:{lst[3]:02d}"
        return formatted

    def read_number_of_tariffs(self):
        """

        :return: 电表number_of_tariffs值
        """
        address = MemoryAddr.number_of_tariffs
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        return measure_value[0]

    def read_number_of_seasons(self):
        """

        :return: 电表number_of_seasons值
        """
        address = MemoryAddr.number_of_seasons
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        return measure_value[0]

    def read_number_of_schedules(self):
        """

        :return: 电表number_of_schedules值
        """
        address = MemoryAddr.number_of_schedules
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        return measure_value[0]

    def read_number_of_segments(self):
        """

        :return: 电表number_of_segments值
        """
        address = MemoryAddr.number_of_segments
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        return measure_value[0]

    def read_segment_1_setting(self):
        """

        :return: 电表read_segment_1_setting值
        """
        address = MemoryAddr.segment_1_setting
        count = 3
        slave = self.slave_id
        lst = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        formatted = f"{lst[0]:02d}:{lst[1]:02d} {lst[2]:01d}"
        return lst, formatted

    def read_device_time(self):
        """
        读设备时间
        @return:
        """
        address = MemoryAddr.device_time
        count = MemoryReg.reg_uint16
        slave = self.slave_id
        lst = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        # formatted = f"{lst[0]:02d}_{lst[1]:02d}_{lst[2]:02d} {lst[3]:02d}:{lst[4]:02d}:{lst[5]:02d}"
        return lst

    def set_device_time(self, year: int = 2000, month: int = 1, day: int = 1, hour: int = 0, minute: int = 0,
                        second: int = 0):
        """
        设置电表时间
        @param year:
        @param month:
        @param day:
        @param hour:
        @param minute:
        @param second:
        @return:
        """
        address = MemoryAddr.device_time
        values = year, month, day, hour, minute, second
        slave = self.slave_id
        count = MemoryReg.reg_uint16
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_device_time fail, ret is:{ret}')
            return False
        return True


if __name__ == '__main__':
    rm = HandleMemory()
    # rm.set_device_time()
    r = rm.read_enable_special_weekday_schedule()
    print(r)
