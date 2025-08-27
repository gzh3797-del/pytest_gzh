import time
from Config.IOM.modbus_set_attr import set_ao_pmi


class TestAo:
    def ttest_single_ao(self, current, expected):
        """
        仅配置AO physical measurement Input为current
        :param current: 输入数值列表
        :param expected: 预期范围
        :return:
        """
        st = time.time()
        for t in range(4):
            time.sleep(6)
            print(f"*************************开始执行AO{t + 1}*************************")
            ao_number = t + 1
            for i in range(len(current)):
                set_ao_pmi(ao_number, current[i])
                print(f"AO{ao_number} 开始配置为{current[i]},预期为{expected[i]}V")
                time.sleep(6)
            if ao_number < 4:
                print(
                    f"*************************AO{ao_number} 配置完成,你有6s时间切换到AO{ao_number + 1}*************************")
            else:
                print(f"*************************AO{ao_number} 配置完成,测试结束*************************")
        et = time.time()
        print(f"配置耗时{et - st}")
        pass
