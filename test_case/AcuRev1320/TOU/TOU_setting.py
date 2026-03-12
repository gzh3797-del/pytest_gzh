#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# @File     :TOU_setting.py
# @Author   :lcs
# @Time     :2026/1/20
# @Desc     :
import time

from comm.source_control import *
from comm.QT_comm.QT_utils.QT_auto_utils_1320 import AutoHelper

from test_case.AcuRev1320.TOU.page_elements_config import PageAddr

from test_case.AcuRev1320.TOU.tou_modbus_get import HandleMemory
from tools.log import Log

Log(str(__file__).split("\\")[-1])
# 升级包路径
package_path_target = r'C:\autotest_local\update_version\AcuRev-4100_Application_v1.01p35_20251126.MFEA'
package_path_base = r'C:\autotest_local\update_version\AcuRev-4100_Application_v1.01p32_20251025.MFEA'

# ================= 配置区 =================
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ACUVIEW_EXE = r'C:\Users\CongsongLiu\Acuview2\Acuview 2.exe'
device_image_path = "page_elements\\AucRev1320\\1320 TCP"


# =========================================

class TestAcuviewTou:
    """Acuview 自动化测试类"""

    def __init__(self):
        """每个测试用例前的准备工作"""
        self.app_path = PageAddr.Acuview_exe
        # self.device_image_path = PageAddr.device_image_path
        self.helper = AutoHelper(confidence=0.8)
        # self.modbus_client = ModbusRtuOrTcp()
        self.handle_memory = HandleMemory(slave_id=1)
        # self.helper.kill_acuview_apps()

    def login(self, connect_mode: int = 1):
        """
        打开上位机，连接登录AcuRev_1320设备
        @param connect_mode: 0:TCP方式 1:RTU方式
        @return:
        """
        if connect_mode == 0:
            device_path = PageAddr.device_connect_TCP
            self.helper.hotkey('win', 'd')
            self.helper.wait(1)
            self.helper.launch_app(self.app_path)
            self.helper.connect_device(device_path)
            time.sleep(2)
            res = self.helper.check_image_exists(PageAddr.login_success)
            # print(res)
            logging.info(f'login:{res}')
            print(f'login:{res}')
            return res
        else:
            device_path = PageAddr.device_connect_RTU
            self.helper.hotkey('win', 'd')
            self.helper.wait(2)
            self.helper.launch_app(self.app_path)
            time.sleep(2)
            self.helper.click_image(device_path, offset_x=-163)
            self.helper.click_pos((1321, 143))
            res = self.helper.check_image_exists(PageAddr.RTU_connect_failed)
            if res:
                self.helper.click_image(PageAddr.yes)
                time.sleep(5)
                self.helper.click_pos((216, 167))
            time.sleep(3)
            check_res = self.helper.check_image_exists(PageAddr.login_success)
            print(res)
            logging.info(f'login:{check_res}')
            print(f'login:{check_res}')
            return check_res

    def enter_tou_setting(self):
        """
        进入上位机tou_setting页面
        @return: 进入上位机tou_setting结果
        """
        self.helper.click_pos(PageAddr.setting)
        self.helper.click_pos(PageAddr.tou_setting)
        self.helper.click_pos(PageAddr.tou)
        time.sleep(1)
        cmp_res = self.helper.check_image_exists(PageAddr.enter_tou_setting_success)
        logging.info(f'进入tou_setting页面:{cmp_res}')
        print(f'进入tou_setting页面:{cmp_res}')
        return cmp_res

    def confirm_password(self):
        """
        进入上位机设置general页面，确认设置密码
        @return:
        """
        self.helper.click_pos(PageAddr.setting)
        self.helper.click_pos(PageAddr.metering)
        self.helper.click_pos(PageAddr.general)
        self.helper.click_pos(PageAddr.password)
        self.helper.click_image(PageAddr.confirm)

    def click_update(self):
        """
        上位机上点击update;
        匹配更新成功图片，检查上位机是否更新成功
        @return:bool
        """
        self.helper.click_image(PageAddr.update)
        self.helper.click_image(PageAddr.yes)
        res = self.helper.check_image_exists(PageAddr.update_successful)
        if res is True:
            self.helper.click_image(PageAddr.yes)
        return res

    def tou_reset(self):
        """
        上位机TOU setting 页面上点击reset to default，恢复TOU设置
        @return:
        """
        time.sleep(1)
        self.helper.click_image(PageAddr.tou_reset)
        self.helper.click_image(PageAddr.yes)
        self.helper.click_image(PageAddr.yes)

    def open_TOU_enable(self):
        """
        TOU_enable使能打开
        @return:
        """
        # 上位机点击右侧进度条，返回TOU页面顶部
        self.helper.click_pos((1710, 257))
        self.helper.click_pos((1710, 257))
        tou_enable = self.handle_memory.read_enable_tou()
        if tou_enable == 0:
            self.helper.click_pos(PageAddr.tou_enable)

    def close_TOU_enable(self):
        """
        TOU_enable使能关闭
        @return:
        """
        # 上位机点击右侧进度条，返回TOU页面顶部
        self.helper.click_pos((1710, 325))
        tou_enable = self.handle_memory.read_enable_tou()
        if tou_enable == 1:
            self.helper.click_pos(PageAddr.tou_enable)

    def billing_and_tariff_Setting(self, billing_mode: int = 0, at_time: str = '01 00:00:00', tariff: int = 0):
        """
        上位机TOU Setting页面设置billing_and_tariff
        @param billing_mode: 0:end_of_month 1:Assign
        @param at_time:上位机TOU Setting页面at_time时间，格式:'01 00:00:00'
        @param tariff: 0:尖sharp 1:峰peak 2:谷valley 3:平normal
        @return:
        """
        if billing_mode == 0:
            self.helper.click_image(PageAddr.end_of_month)

        if billing_mode == 1:
            self.helper.click_image(PageAddr.Assign)
        self.helper.click_image(PageAddr.at_time, offset_x=80)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(at_time)
        if tariff == 0:
            self.helper.click_image(PageAddr.tariff_sharp)
        elif tariff == 1:
            self.helper.click_image(PageAddr.tariff_sharp)
            self.helper.click_image(PageAddr.tariff_peak)
        elif tariff == 2:
            self.helper.click_image(PageAddr.tariff_sharp)
            self.helper.click_image(PageAddr.tariff_peak)
            self.helper.click_image(PageAddr.tariff_valley)
        elif tariff == 3:
            self.helper.click_image(PageAddr.tariff_sharp)
            self.helper.click_image(PageAddr.tariff_peak)
            self.helper.click_image(PageAddr.tariff_valley)
            self.helper.click_image(PageAddr.tariff_normal)

    def tou_schedules_add(self, schedule_id: int = 1, segment_id: int = 1,
                          segment_time: list = ['00:00', '01:00', '02:00', '03:00', '04:00', '05:00', '06:00',
                                                '07:00', '08:00', '09:00', '12:00', '13:00', '14:00', '15:00'],
                          segment_tariff: list = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]):
        """
        上位机TOU Setting页面增加schedules
        @param schedule_id: 需要增加都少个schedules，此处填写多少
        @param segment_id: 每一个schedules需要增加多少个segment，此处填写多少
        @param segment_time: 每一个segment的segment_time
        @param segment_tariff: 每一个segment的segment_tariff
        @return:
        """
        # 页面Y轴滚动条向下点击15次
        for i in range(15):
            self.helper.click_pos((1710, 914))
        (edit_x, edit_y) = PageAddr.tou_schedules_edit
        (segment_time_x, segment_time_y) = PageAddr.segment_time
        for i in range(schedule_id):
            self.helper.click_pos(PageAddr.tou_schedules_add)
            self.helper.click_pos((edit_x, edit_y))
            edit_y += 30
            for j in range(segment_id):
                if j != 0:
                    self.helper.click_pos(PageAddr.segment_add)
                self.helper.click_pos((segment_time_x, segment_time_y))
                self.helper.hotkey('ctrl', 'a')
                self.helper.paste_text(segment_time[j])
                self.helper.click_pos((segment_time_x + 110, segment_time_y))
                k = (segment_tariff[j] + 1) * 19
                self.helper.click_pos((segment_time_x + 110, segment_time_y + k))
                segment_time_y += 30
            segment_time_y = segment_time_y - 30 * segment_id
            self.helper.click_image(PageAddr.confirm)

    def check_segment_tariff(self, segment_tariff: int = 0):
        """
        上检查segment ID费率是否只能选择“sharp、peak、valley、normal”;
        当tariff设置为0时，上位机segment_tariff只能选择sharp;
        当tariff设置为1时，上位机segment_tariff只能选择sharp、peak;
        当tariff设置为3时，上位机segment_tariff只能选择sharp、peak、valley、normal;
        @param segment_tariff: 0:尖sharp 1:峰peak 2:谷valley 3:平normal;
        @return:检查结果：true or false
        """
        if segment_tariff == 0:
            (segment_time_x, segment_time_y) = PageAddr.segment_time
            self.helper.click_pos(PageAddr.tou_schedules_edit)
            self.helper.click_pos((segment_time_x + 110, segment_time_y))
            result = self.helper.check_image_exists(PageAddr.check_tariff_sharp)
            time.sleep(1)
            self.helper.double_click_image(PageAddr.tou_schedule, offset_x=330)
            return result
        if segment_tariff == 1:
            (segment_time_x, segment_time_y) = PageAddr.segment_time
            self.helper.click_pos(PageAddr.tou_schedules_edit)
            self.helper.click_pos((segment_time_x + 110, segment_time_y))
            result = self.helper.check_image_exists(PageAddr.check_tariff_sharp_peak)
            time.sleep(1)
            self.helper.double_click_image(PageAddr.tou_schedule, offset_x=330)
            return result
        if segment_tariff == 3:
            (segment_time_x, segment_time_y) = PageAddr.segment_time
            self.helper.click_pos(PageAddr.tou_schedules_edit)
            self.helper.click_pos((segment_time_x + 110, segment_time_y))
            result = self.helper.check_image_exists(PageAddr.check_tariff_sharp_peak_valley_normal)
            time.sleep(1)
            self.helper.double_click_image(PageAddr.tou_schedule, offset_x=330)
            return result

    def tou_schedules_add_15(self):
        """
        配置第15个schedules失败，检查上位机是否报错the limit of index
        @return: result
        """
        self.helper.click_pos(PageAddr.tou_schedules_edit)
        self.helper.click_pos(PageAddr.segment_add)
        result = self.helper.check_image_exists(PageAddr.the_limit_of_index)
        self.helper.click_image(PageAddr.yes)
        self.helper.click_image(PageAddr.tou_schedule, offset_x=330)
        return result

    def tou_seasons_add(self, session_id: int = 1,
                        start_date: list = ['01-01', '02-01', '03-01', '04-01', '05-01', '06-01',
                                            '07-01', '08-01', '09-01', '10-01', '11-01', '12-01'],
                        seasons_schedule_id: list = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]):
        """
        上位机TOU Setting页面增加seasons
        @param session_id: 需要增加都少个seasons，此处填写多少
        @param start_date: 每一个seasons的start_date
        @param seasons_schedule_id: 每一个seasons对应的schedule_id
        @return:
        """
        (start_date_x, start_date_y) = PageAddr.start_date
        (schedule_id_x, schedule_id_y) = PageAddr.sessions_schedule_id
        time.sleep(2)
        for i in range(session_id):
            self.helper.click_pos(PageAddr.tou_sessions_add)
            self.helper.click_pos((start_date_x, start_date_y))
            self.helper.hotkey('ctrl', 'a')
            self.helper.paste_text(start_date[i])
            start_date_y += 30
            j = 0
            if seasons_schedule_id[i] <= 9:
                self.helper.click_pos((schedule_id_x, schedule_id_y))
                j = (seasons_schedule_id[i] - 1) * 16
                self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))
            else:
                self.helper.click_pos((schedule_id_x, schedule_id_y))
                j = (9 - 1) * 16
                if seasons_schedule_id[i] == 10:
                    self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                if seasons_schedule_id[i] == 11:
                    self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                if seasons_schedule_id[i] == 12:
                    self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))
            schedule_id_y += 30

    def tou_seasons_remove(self, number: int = 1):
        """
        上位机上删除seasons
        @param number: 需要删除的seasons个数
        @return:
        """
        (sessions_remove_x, sessions_remove_y) = PageAddr.sessions_remove
        for i in range(number):
            self.helper.click_pos((sessions_remove_x, sessions_remove_y))

    def tou_schedules_remove(self, number: int = 1):
        """
        上位机上删除schedules
        @param number: 需要删除的schedules个数
        @return:
        """
        for i in range(number):
            self.helper.click_image(PageAddr.schedules_remove, offset_x=60, index=number)
            number -= 1

    def tou_segment_remove(self, number: int = 1):
        """
        上位机上删除segment
        @param number: 需要删除的segment个数
        @return:
        """
        self.helper.click_pos(PageAddr.tou_schedules_edit)
        for i in range(number):
            self.helper.click_image(PageAddr.segment_remove, offset_x=31, index=number)
            number -= 1
        self.helper.click_image(PageAddr.confirm)

    def tou_seasons_add_13(self):
        """
        配置第15个schedules失败，检查上位机是否报错the limit of index
        @return: result
        """
        self.helper.click_pos(PageAddr.tou_sessions_add)
        result = self.helper.check_image_exists(PageAddr.the_limit_of_index)
        self.helper.click_image(PageAddr.yes)
        return result

    def tou_weekends_add(self, weekend_selection: list = [0], weekend_schedule: int = 0):
        """
        TOU Weekends栏，Weekend Selection勾选对应星期，weekend_schedule选择schedule id
        @param weekend_selection: list类型，其中：0表示周一；1到周二；2表示周三；3到周四；5表示周六；6到周日
        @param weekend_schedule: weekend_schedule选择对应schedule id
        @return:
        """
        # 上位机点击右侧进度条，返回TOU页面底部
        self.helper.click_pos((1710, 900))
        self.helper.click_pos((1710, 900))
        (x, y) = PageAddr.weekend_schedule_id
        for i in weekend_selection:
            self.helper.click_image(PageAddr.tou_weekend_selection_Mon, offset_x=(70 * i))
        if weekend_schedule <= 9:
            self.helper.click_pos((x, y))
            j = (weekend_schedule - 1) * 16
            self.helper.click_pos((x, y + 40 + j))
        if weekend_schedule > 9:
            self.helper.click_pos((x, y))
            for k in range(weekend_schedule - 9):
                self.helper.click_pos((x + 50, y + 168))
            self.helper.click_pos((x, y + 168))

    def open_enable_special_weekday_schedule(self):
        """
        enable_special_weekday_schedule使能打开
        @return:
        """
        # 上位机点击右侧进度条，返回TOU页面顶部
        self.helper.click_pos((1710, 900))
        self.helper.click_pos((1710, 900))
        tou_enable = self.handle_memory.read_enable_special_weekday_schedule()
        if tou_enable == 0:
            self.helper.click_pos(PageAddr.enable_special_weekday_schedule)

    def open_holiday_setting_enable(self):
        """
        holiday_setting_enable使能打开
        @return:
        """

        tou_enable = self.handle_memory.read_holiday_setting_enable()
        if tou_enable == 0:
            self.helper.click_pos(PageAddr.setting)
            self.helper.click_pos(PageAddr.tou_setting)
            self.helper.click_pos(PageAddr.ten_years_holiday)
            # 上位机点击右侧进度条，返回TOU页面顶部
            self.helper.click_pos((1710, 325))
            self.helper.click_pos(PageAddr.holiday_setting_enable)

    def close_enable_special_weekday_schedule(self):
        """
        enable_special_weekday_schedule使能关闭
        @return:
        """
        # 上位机点击右侧进度条，返回TOU页面顶部
        self.helper.click_pos((1710, 900))
        self.helper.click_pos((1710, 900))
        tou_enable = self.handle_memory.read_enable_special_weekday_schedule()
        if tou_enable == 1:
            self.helper.click_pos(PageAddr.enable_special_weekday_schedule)

    def close_holiday_setting_enable(self):
        """
        holiday_setting_enable使能关闭
        @return:
        """

        tou_enable = self.handle_memory.read_holiday_setting_enable()
        if tou_enable == 1:
            self.helper.click_pos(PageAddr.setting)
            self.helper.click_pos(PageAddr.tou_setting)
            self.helper.click_pos(PageAddr.ten_years_holiday)
            # 上位机点击右侧进度条，返回TOU页面顶部
            self.helper.click_pos((1710, 325))
            self.helper.click_pos(PageAddr.holiday_setting_enable)

    def tou_special_weekday_schedule_add(self, session_id: int = 1,
                                         weekend_selection=[[0], [0], [0], [0], [0], [0], [0], [0], [0], [0],
                                                            [0], [0]],
                                         weekend_schedule=[[1], [1], [1], [1], [1], [1], [1], [1], [1], [1],
                                                           [1], [1]]):
        """
        TOU 设置页面添加special_weekday_schedule
        @param session_id: 添加多少个session_id，这里填写多少
        @param weekend_selection: list类型，个其中每个子list对应每个session_id包含多少个星期，其中：0表示周一；1到周二；2表示周三；3到周四；5表示周六；6到周日
        @param weekend_schedule: list类型，个其中每个子list对应每个session_id中每个星期的session
        @return:
        """
        # 上位机点击右侧进度条，返回TOU页面底部
        self.helper.click_pos((1710, 900))
        self.helper.click_pos((1710, 900))
        (edit_x, edit_y) = PageAddr.weekday_schedule_edit
        (x, y) = PageAddr.special_weekday_selection_Mon
        for i in range(session_id):
            self.helper.click_pos(PageAddr.special_weekday_schedule_add)
            self.helper.click_pos((edit_x, edit_y))
            edit_y += 30
            for j in range(len(weekend_selection[i])):
                self.helper.click_pos((x, y + (40 * weekend_selection[i][j])))
                self.helper.click_pos((x + 140, y + (40 * weekend_selection[i][j])))
                self.helper.click_pos(
                    (x + 140, y + (40 * weekend_selection[i][j] + 40 + 16 * (weekend_schedule[i][j] - 1))))
            self.helper.click_image(PageAddr.weekend_schedule_confirm)

    def tou_holidays_add(self, holidays_id: int = 1,
                         start_date: list = ['01-01', '01-10', '01-20', '02-01', '02-10', '02-20',
                                             '03-01', '03-10', '03-20', '04-01', '04-10', '04-20',
                                             '05-01', '05-10', '05-20', '06-01', '06-10', '06-20',
                                             '07-01', '07-10', '07-20', '08-01', '08-10', '08-20',
                                             '09-01', '09-10', '09-20', '10-01', '11-01', '12-01'],
                         schedule_id: list = [1, 1, 1, 1, 1, 1,
                                              1, 1, 1, 1, 1, 1,
                                              1, 1, 1, 1, 1, 1,
                                              1, 1, 1, 1, 1, 1,
                                              1, 1, 1, 1, 1, 1]):
        """
        上位机TOU Setting页面增加tou_holidays
        @param holidays_id: 需要增加都少个tou_holidays，此处填写多少
        @param start_date: 每一个tou_holidays的start_date
        @param schedule_id: 每一个tou_holidays对应的schedule_id
        @return:
        """
        # 上位机点击右侧进度条，返回TOU页面底部
        self.helper.click_pos((1710, 900))
        self.helper.click_pos((1710, 900))
        (start_date_x, start_date_y) = PageAddr.tou_holidays_start_date
        (schedule_id_x, schedule_id_y) = PageAddr.tou_holidays_schedule_id
        time.sleep(2)
        for i in range(holidays_id):
            if i < 16:
                self.helper.click_pos(PageAddr.tou_holidays_add)
                self.helper.click_pos((start_date_x, start_date_y))
                self.helper.hotkey('ctrl', 'a')
                self.helper.paste_text(start_date[i])
                start_date_y += 30
                j = 0
                if schedule_id[i] <= 9:
                    self.helper.click_pos((schedule_id_x, schedule_id_y))
                    j = (schedule_id[i] - 1) * 16
                    self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))
                else:
                    self.helper.click_pos((schedule_id_x, schedule_id_y))
                    j = (9 - 1) * 16
                    if schedule_id[i] == 10:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    if schedule_id[i] == 11:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    if schedule_id[i] == 12:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))
                schedule_id_y += 30
            else:
                self.helper.click_pos(PageAddr.tou_holidays_add)
                self.helper.click_pos((953, 841))
                self.helper.click_pos((643, 833))
                self.helper.hotkey('ctrl', 'a')
                self.helper.paste_text(start_date[i])
                schedule_id_x = 770
                schedule_id_y = 840
                if schedule_id[i] <= 9:
                    self.helper.click_pos((schedule_id_x, schedule_id_y))
                    j = (schedule_id[i] - 1) * 16
                    self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))
                else:
                    self.helper.click_pos((schedule_id_x, schedule_id_y))
                    j = (9 - 1) * 16
                    if schedule_id[i] == 10:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    if schedule_id[i] == 11:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    if schedule_id[i] == 12:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))

    def year1_holidays_add(self, holidays_id: int = 1,
                         start_date: list = ['01-01', '01-10', '01-20', '02-01', '02-10', '02-20',
                                             '03-01', '03-10', '03-20', '04-01', '04-10', '04-20',
                                             '05-01', '05-10', '05-20', '06-01', '06-10', '06-20',
                                             '07-01', '07-10', '07-20', '08-01', '08-10', '08-20',
                                             '09-01', '09-10', '09-20', '10-01', '11-01', '12-01'],
                         schedule_id: list = [1, 1, 1, 1, 1, 1,
                                              1, 1, 1, 1, 1, 1,
                                              1, 1, 1, 1, 1, 1,
                                              1, 1, 1, 1, 1, 1,
                                              1, 1, 1, 1, 1, 1]):
        """
        上位机TOU Setting页面增加tou_holidays
        @param holidays_id: 需要增加都少个tou_holidays，此处填写多少
        @param start_date: 每一个tou_holidays的start_date
        @param schedule_id: 每一个tou_holidays对应的schedule_id
        @return:
        """
        (start_date_x, start_date_y) = PageAddr.year1_holidays_start_date
        (schedule_id_x, schedule_id_y) = PageAddr.year1_holidays_schedule_id
        time.sleep(2)
        for i in range(holidays_id):
            if i < 16:
                self.helper.click_pos(PageAddr.year1_holidays_add)
                self.helper.click_pos((start_date_x, start_date_y))
                self.helper.hotkey('ctrl', 'a')
                self.helper.paste_text(start_date[i])
                start_date_y += 30
                j = 0
                if schedule_id[i] <= 9:
                    self.helper.click_pos((schedule_id_x, schedule_id_y))
                    j = (schedule_id[i] - 1) * 16
                    self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))
                else:
                    self.helper.click_pos((schedule_id_x, schedule_id_y))
                    j = (9 - 1) * 16
                    if schedule_id[i] == 10:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    if schedule_id[i] == 11:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    if schedule_id[i] == 12:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))
                schedule_id_y += 30
            else:
                self.helper.click_pos(PageAddr.tou_holidays_add)
                self.helper.click_pos((953, 913))
                self.helper.click_pos((645, 909))
                self.helper.hotkey('ctrl', 'a')
                self.helper.paste_text(start_date[i])
                schedule_id_x = 769
                schedule_id_y = 910
                if schedule_id[i] <= 9:
                    self.helper.click_pos((schedule_id_x, schedule_id_y))
                    j = (schedule_id[i] - 1) * 16
                    self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))
                else:
                    self.helper.click_pos((schedule_id_x, schedule_id_y))
                    j = (9 - 1) * 16
                    if schedule_id[i] == 10:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    if schedule_id[i] == 11:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    if schedule_id[i] == 12:
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                        self.helper.click_pos((schedule_id_x + 56, schedule_id_y + 40 + j))
                    self.helper.click_pos((schedule_id_x, schedule_id_y + 40 + j))
    def tou_holidays_add_31(self):
        """
        添加第31个tou_holidays
        @return:
        """
        self.helper.click_pos(PageAddr.tou_holidays_add)
        result = self.helper.check_image_exists(PageAddr.the_limit_of_index)
        if result is True:
            self.helper.click_image(PageAddr.yes)
        return result

    def tou_special_weekday_schedule_remove(self, number: int = 1):
        """
        上位机上删除special weekday schedules
        @param number: 需要删除的schedules个数
        @return:
        """
        # 上位机点击右侧进度条，返回TOU页面底部
        self.helper.click_pos((1710, 900))
        self.helper.click_pos((1710, 900))
        for i in range(number):
            self.helper.click_image(PageAddr.schedules_remove, offset_x=60, index=number)
            number -= 1

    def tou_holidays_remove(self, number: int = 1):
        """
        上位机上删除tou_holidays
        @param number: 需要删除的tou_holidays个数
        @return:
        """
        # 上位机点击右侧进度条，返回TOU页面底部
        self.helper.click_pos((1710, 900))
        self.helper.click_pos((1710, 900))
        for i in range(number):
            self.helper.click_image(PageAddr.remove)
            number -= 1

    def tou_year1_holidays_remove(self, number: int = 1):
        """
        上位机上删除special weekday schedules
        @param number: 需要删除的schedules个数
        @return:
        """
        for i in range(number):
            self.helper.click_image(PageAddr.remove, index=number)
            number -= 1

    def check_season_ID_undefined(self):
        """
        Special Weekday Schedule栏，配置添加1个不存在的season，检查是否出现season ID undefined弹窗
        @return:
        """
        self.helper.click_image(PageAddr.update)
        result = self.helper.check_image_exists(PageAddr.season_ID_undefined)
        if result is True:
            self.helper.click_image(PageAddr.yes)
        return result

    def tou_creation(self, billing_mode: int = 0, at_time: str = '01 00:00:00', tariff: int = 0, session_id: int = 1,
                     start_date: list = ['01-01'], seasons_schedule_id: list = [1], schedule_id: int = 1,
                     segment_id: int = 1, segment_time: str = '00:00',
                     segment_tariff: int = 0):
        self.open_TOU_enable()
        self.billing_and_tariff_Setting(billing_mode=billing_mode, at_time=at_time, tariff=tariff)
        self.tou_schedules_add(schedule_id=schedule_id)
        self.tou_seasons_add(session_id=session_id, start_date=start_date, seasons_schedule_id=seasons_schedule_id)
        self.click_update()

    def check_open_TOU_enable(self):
        enable_tou = self.handle_memory.read_enable_tou()
        if enable_tou == 1:
            logging.debug(f'open_TOU_enable成功')
            print(f'open_TOU_enable成功')
            return True
        else:
            logging.info(f'open_TOU_enable失败')
            print(f'open_TOU_enable失败')
            return False

    def check_billing_mode(self, billing_mode: int = 0):
        monthly_billing_mode = self.handle_memory.read_monthly_billing_mode()
        if monthly_billing_mode == billing_mode:
            logging.info(f'设置monthly_billing_mode:{billing_mode}成功')
            print(f'设置monthly_billing_mode:{billing_mode}成功')
            return True
        else:
            logging.info(f'设置monthly_billing_mode:{billing_mode}失败')
            print(f'设置monthly_billing_mode:{billing_mode}失败')
            return False

    def check_at_time(self, at_time: str = '01 00:00:00'):
        billing_time = self.handle_memory.read_billing_time()
        if billing_time == at_time:
            logging.info(f'设置billing_time:{at_time}成功')
            print(f'设置billing_time:{at_time}成功')
            return True
        else:
            logging.info(f'设置billing_time:{at_time}失败')
            print(f'设置billing_time:{at_time}失败')
            return False

    def check_number_of_tariffs(self, tariff: int = 0):
        number_of_tariffs = self.handle_memory.read_number_of_tariffs()
        if number_of_tariffs == tariff:
            logging.info(f'设置number_of_tariffs:{tariff}成功')
            print(f'设置number_of_tariffs:{tariff}成功')
            return True
        else:
            logging.info(f'设置number_of_tariffs:{tariff}失败')
            print(f'设置number_of_tariffs:{tariff}失败')
            return False

    def check_segment_1_setting_tariffs(self, tariff):
        (lst, _) = self.handle_memory.read_segment_1_setting()
        if lst[2] == tariff:
            logging.info(f'检查设置segment_1_setting_tariffs:{tariff}成功')
            print(f'检查设置segment_1_setting_tariffs:{tariff}成功')
            return True
        else:
            logging.info(f'检查设置segment_1_setting_tariffs:{tariff}失败')
            print(f'检查设置segment_1_setting_tariffs:{tariff}失败')
            return False

    def check_number_of_segments(self, number: int = 0):
        """
        检查segments个数：读取电表寄存器segments个数，传参期望的segments个数，返回bool
        @param number: 期望的segments个数
        @return: bool
        """
        number_of_tariffs = self.handle_memory.read_number_of_segments()
        if number_of_tariffs == number:
            logging.info(f'设置number_of_segments:{number}成功')
            print(f'设置number_of_segments:{number}成功')
            return True
        else:
            logging.info(f'设置number_of_segments:{number}失败')
            print(f'设置number_of_segments:{number}失败')
            return False

    def check_number_of_seasons(self, number: int = 0):
        """
        检查seasons个数：读取电表寄存器seasons个数，传参期望的seasons个数，返回bool
        @param number: 期望的seasons个数
        @return: bool
        """
        number_of_tariffs = self.handle_memory.read_number_of_seasons()
        if number_of_tariffs == number:
            logging.info(f'设置number_of_seasons:{number}成功')
            print(f'设置number_of_seasons:{number}成功')
            return True
        else:
            logging.info(f'设置number_of_seasons:{number}失败')
            print(f'设置number_of_seasons:{number}失败')
            return False

    def check_number_of_schedules(self, number: int = 0):
        number_of_tariffs = self.handle_memory.read_number_of_schedules()
        if number_of_tariffs == number:
            logging.info(f'设置number_of_schedules:{number}成功')
            print(f'设置number_of_schedules:{number}成功')
            return True
        else:
            logging.info(f'设置number_of_schedules:{number}失败')
            print(f'设置number_of_schedules:{number}失败')
            return False

    def read_dst_enable(self):
        enable_tou = self.handle_memory.read_dst_enable()
        if enable_tou == 1:
            logging.debug(f'open_dst_enable成功')
            print(f'open_TOU_enable成功')
            return True
        else:
            logging.info(f'open_dst_enable失败')
            print(f'open_TOU_enable失败')
            return False

    def enter_DST_setting(self):
        """
        进入上位机dst_setting页面
        @return: 进入上位机dst_setting结果
        """
        self.helper.click_pos(PageAddr.setting)
        self.helper.click_image(PageAddr.dst_setting)
        time.sleep(1)
        cmp_res = self.helper.check_image_exists(PageAddr.enter_dst_setting_success)
        logging.info(f'进入dst_setting页面:{cmp_res}')
        print(f'进入dst_setting页面:{cmp_res}')
        return cmp_res

    def open_DST_enable(self):
        """
        dst_enable使能打开
        @return:
        """
        tou_enable = self.handle_memory.read_dst_enable()
        if tou_enable == 0:
            self.helper.click_pos(PageAddr.dst_enable)

    def close_DST_enable(self):
        """
        dst_enable使能关闭
        @return:
        """
        tou_enable = self.handle_memory.read_dst_enable()
        if tou_enable == 1:
            self.helper.click_pos(PageAddr.dst_enable)
            return True
        else:
            return False

    def DST_format(self, dst_format: int = 0):
        """
        上位机设置DST_format；0: format 1(fixed date) 1: format 2 (non fixed date)
        @return:
        """
        if dst_format == 0:
            self.helper.click_image(PageAddr.DST_formal, offset_x=100)
            self.helper.click_image(PageAddr.DST_formal, offset_x=100, offset_y=20)
        else:
            self.helper.click_image(PageAddr.DST_formal, offset_x=100)
            self.helper.click_image(PageAddr.DST_formal, offset_x=100, offset_y=40)

    def setting_format1_DST_time(self, start_time: str = '03-01 00:00', start_adjust_time: int = 60,
                                 end_time: str = '11-01 00:00', end_adjust_time: int = 60):
        """
        上位机设置format1_DST_time和adjust_time
        @param start_time: format1_DST_start_time
        @param start_adjust_time: format1_DST_start_adjust_time
        @param end_time: format1_DST_end_time
        @param end_adjust_time: format1_DST_end_adjust_time
        @return:
        """
        self.helper.click_pos(PageAddr.format1_start_time)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(start_time)
        self.helper.click_pos(PageAddr.format1_start_adjust_time)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(start_adjust_time)

        self.helper.click_pos(PageAddr.format1_end_time)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(end_time)
        self.helper.click_pos(PageAddr.format1_end_adjust_time)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(end_adjust_time)

    def check_open_DST_enable(self):
        """
        检查DST_enable是否开启
        @return:
        """
        enable_tou = self.handle_memory.read_dst_enable()
        if enable_tou == 1:
            logging.debug(f'open_DST_enable成功')
            print(f'open_DST_enable成功')
            return True
        else:
            logging.info(f'open_DST_enable失败')
            print(f'open_DST_enable失败')
            return False

    def check_device_time(self):
        """
        检查设备时间
        @return:
        """
        (year, month, day, hour, minute, second) = self.handle_memory.read_device_time()
        return year, month, day, hour, minute, second

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
        self.handle_memory.set_device_time(year, month, day, hour, minute, second)

    def check_format_adjust_time_invalid(self):
        """
        检查format_adjust_time_invalid
        @return:
        """
        self.helper.click_image(PageAddr.update)
        res = self.helper.check_image_exists(PageAddr.adjust_time_invalid)
        self.helper.click_image(PageAddr.yes)
        return res

    def setting_format2_DST_time(self, start_month: int = 1, start_day: int = 1, start_week: int = 1,
                                 start_time: str = '00:00', start_adjust_time: int = 60,
                                 end_month: int = 1, end_day: int = 1, end_week: int = 1,
                                 end_time: str = '00:00', end_adjust_time: int = 60):
        """
        上位机设置format2_DST_time和adjust_time
        @param start_month:开始月份，可选择1-12月，对应1月到12月
        @param start_day:开始周数，可选择第1-5周
        @param start_week:开始星期，可选择第1-7，对应周日到周六 1:周日 2:周一 3:周二 ··· 7:周六
        @param start_time:开始时间，格式'00:00'
        @param start_adjust_time:调整时间，单位min
        @param end_month:结束月份，可选择1-12月，对应1月到12月
        @param end_day:结束周数，可选择第1-5周
        @param end_week:结束星期，可选择第1-7，对应周日到周六 1:周日 2:周一 3:周二 ··· 7:周六
        @param end_time:结束时间，格式'00:00'
        @param end_adjust_time:结束调整时间，单位min
        @return:
        """
        (start_month_x, start_month_y) = PageAddr.format2_start_month
        if start_month <= 10:
            self.helper.click_pos((start_month_x, start_month_y))
            self.helper.click_pos((start_month_x + 63, start_month_y + 18))
            self.helper.click_pos((start_month_x + 63, start_month_y + 18))
            self.helper.click_pos((start_month_x, start_month_y + (18 * start_month)))
        elif start_month == 11:
            self.helper.click_pos((start_month_x, start_month_y))
            self.helper.click_pos((start_month_x + 63, start_month_y + 180))
            self.helper.click_pos((start_month_x, start_month_y + 180))
        else:
            self.helper.click_pos((start_month_x, start_month_y))
            self.helper.click_pos((start_month_x + 63, start_month_y + 180))
            self.helper.click_pos((start_month_x + 63, start_month_y + 180))
            self.helper.click_pos((start_month_x, start_month_y + 180))
        (start_day_x, start_day_y) = PageAddr.format2_start_day
        self.helper.click_pos((start_day_x, start_day_y))
        self.helper.click_pos((start_day_x, start_day_y + (18 * start_day)))
        (start_week_x, start_week_y) = PageAddr.format2_start_week
        self.helper.click_pos((start_week_x, start_week_y))
        self.helper.click_pos((start_week_x, start_week_y + (18 * start_week)))
        self.helper.click_pos(PageAddr.format2_start_time)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(start_time)
        self.helper.click_pos(PageAddr.format2_start_adjust_time)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(start_adjust_time)

        (end_month_x, end_month_y) = PageAddr.format2_end_month
        if end_month <= 10:
            self.helper.click_pos((end_month_x, end_month_y))
            self.helper.click_pos((end_month_x + 63, end_month_y + 18))
            self.helper.click_pos((end_month_x + 63, end_month_y + 18))
            self.helper.click_pos((end_month_x, end_month_y + (18 * end_month)))
        elif end_month == 11:
            self.helper.click_pos((end_month_x, end_month_y))
            self.helper.click_pos((end_month_x + 63, end_month_y + 180))
            self.helper.click_pos((end_month_x, end_month_y + 180))
        else:
            self.helper.click_pos((end_month_x, end_month_y))
            self.helper.click_pos((end_month_x + 63, end_month_y + 180))
            self.helper.click_pos((end_month_x + 63, end_month_y + 180))
            self.helper.click_pos((end_month_x, end_month_y + 180))
        (end_day_x, end_day_y) = PageAddr.format2_end_day
        self.helper.click_pos((end_day_x, end_day_y))
        self.helper.click_pos((end_day_x, end_day_y + (18 * end_day)))
        (end_week_x, end_week_y) = PageAddr.format2_end_week
        self.helper.click_pos((end_week_x, end_week_y))
        self.helper.click_pos((end_week_x, end_week_y + (18 * end_week)))
        self.helper.click_pos(PageAddr.format2_end_time)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(end_time)
        self.helper.click_pos(PageAddr.format2_end_adjust_time)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(end_adjust_time)


if __name__ == "__main__":
    TestAcuviewTou = TestAcuviewTou()
    # TestAcuviewTou.login()
    # cmp_res = TestAcuviewTou.enter_tou_setting()
    # TestAcuviewTou.open_TOU_enable()
    # TestAcuviewTou.billing_and_tariff_Setting(billing_mode=0, at_time='01 00:30:00', tariff=3)
    # TestAcuviewTou.tou_schedules_add(schedule_id=1, segment_id=1, segment_time=['01:00'],
    #                                  segment_tariff=[1])
    # TestAcuviewTou.tou_seasons_add(session_id=2, start_date=['01-01', '02-01'], seasons_schedule_id=[1, 2])
    # TestAcuviewTou.click_update()
    # res = TestAcuviewTou.check_tou_setting(at_time='01 00:30:00', tariff=1)
    # print(res)
    # time.sleep(2)
    # TestAcuviewTou.tou_reset()
    # TestAcuviewTou.handle_memory.modbus_client.close()
    # TestAcuviewTou.confirm_password()
    # TestAcuviewTou.enter_DST_setting()
    # TestAcuviewTou.open_DST_enable()
    # TestAcuviewTou.DST_format(dst_format=2)
    # TestAcuviewTou.setting_format1_DST_start(start_time="04-01 00:00", start_adjust_time="50")
    # TestAcuviewTou.click_update()
    # TestAcuviewTou.set_device_time()
    result = TestAcuviewTou.tou_holidays_add_31()
    print(result)
