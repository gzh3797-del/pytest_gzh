import time
import os
import math
from datetime import datetime
# import pytest

# def generate_demand_case(
#         angle_value,
#         current_value,
#         wire_type,
#         demand_method,
#         demand_interval,
#         demand_update_rate
#     ):
#     print(angle_value)
#     print(current_value)
#     print(wire_type)
#     print(demand_method)
#     print(demand_interval)
#     print(demand_update_rate)
#
#
# input_values = [
#     [30,4,4,0,4,10],
#     [30,4,4,0,5,10],
#     [30,4,4,0,4,10],
#     [30,4,4,0,4,10],
#     [30,4,4,0,4,10],
#     [30,4,4,0,4,10],
# ]
#
# @pytest.mark.parametrize(
#     "angle_value,current_value,wire_type,demand_method,demand_interval,demand_update_rate",
#     input_values
# )
# def test_fixed_interval_updaterate(
#     angle_value,
#     current_value,
#     wire_type,
#     demand_method,
#     demand_interval,
#     demand_update_rate
# ):
#     generate_demand_case(
#         angle_value,
#         current_value,
#         wire_type,
#         demand_method,
#         demand_interval,
#         demand_update_rate
#     )



#
# if __name__ == "__main__":
#     # --capture=no 确保 print 在调试模式下可见
#     # -v 输出详细信息
#     # --html=report.html 生成 HTML 报告
#     # pytest.main([__file__, "-v", "--capture=no", "--html=report.html", "--self-contained-html"])
# from datetime import datetime
# import math
#
#
# def calc_wait_seconds(demand_interval_min: int) -> int:
#     now = datetime.now()
#     print(now)
#     seconds_today = now.hour * 3600 + now.minute * 60 + now.second
#     interval_seconds = demand_interval_min * 60
#     next_boundary = math.ceil(seconds_today / interval_seconds) * interval_seconds
#     wait_seconds = next_boundary - seconds_today
#     return wait_seconds if wait_seconds != 0 else interval_seconds


if __name__ == "__main__":
    # wait = calc_wait_seconds(20)
    # print(f"等待: {wait}s")
    # minutes = wait // 60
    # seconds = wait % 60
    # print(f"等待时间: {minutes} 分 {seconds} 秒")

    save_filedir = rf"./precision_measure_{time.strftime('%Y%m%d')}"
    print(save_filedir)
    if not os.path.exists(save_filedir):
        os.makedirs(save_filedir, exist_ok=True)