from comm.AcuRev4100_modbus_get_attr import *
from tools.log import Log
from comm.source_control import *


Log(str(__file__).split("\\")[-1])

Frelist = [50, 60]


def setup_function():
    logging.info('Start：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))


def test_frequency_0x2000():
    rtu_result = []
    for i in Frelist:
        # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 1, 1, 1, i)
        fre = read_frequency(i, 10)
        scale = abs(fre - i) / i
        if i * (1 - 0.0002) <= fre <= i * (1 + 0.0002):
            logging.info(f'源输入频率：{i}，电表读取频率：{fre},实际精度{scale:.3%}')
            rtu_result.append(True)
        else:
            logging.error(f'源输入频率：{i}，电表读取频率：{fre},实际精度{scale:.3%}')
            rtu_result.append(False)
    assert rtu_result == [True, True]


def teardown_function():
    logging.info('End：{}'.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    pass
