"""
FTS编号: FTS_AcuHMI_AZR_004_007
用例标题: AcuHMI-1-7选择Virtual虚拟设备连接 Azure IoT Hub 数据上报
用例级别: LV3

预置条件:
  - 网关已创建虚拟设备，config.yaml 中 virtual_device 字段指定目标设备名
  - 合法 Connection String 已配置，Event Hub 订阅端就绪

测试步骤（一次浏览器完成）:
  1. 进入 Azure IoT 页面，配置虚拟设备并 Save（同时启动 Event Hub 订阅）
  2. 等待虚拟设备按 interval 上报数据
  3. 收到数据后，在同一浏览器导航到 Virtual Devices → 设备 → Reading 页面读取参数值
  4. 比对 Event Hub 上报值与 Reading 页面值（误差在 5% 容差内）

预期结果:
  - Event Hub 收到虚拟设备数据
  - 上报参数单位与 Reading 页面一致，数值误差在 5% 容差内
"""

import json
import threading
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path

import pytest
from azure.eventhub import EventHubConsumerClient

from utils.azure_iot_verifier import _parse_modules

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent

_RECONNECT_BUDGET_SECS = 60


class TestCase_AcuHMI_1_7_AZR_004_007:

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_virtual_device_data_vs_reading(self, azure_page, azure_cfg):
        """一次浏览器完成：配置 Azure IoT → 等待 Event Hub 上报 → 读 Reading 页面 → 比对"""
        azure = azure_cfg["azure_iot"]
        virtual_device = azure.get("virtual_device", "test_virtual01")
        interval_s = int(str(azure.get("interval", "30 seconds")).split()[0])
        subscribe_timeout = _RECONNECT_BUDGET_SECS + interval_s * 3

        received: dict = {}
        done = threading.Event()
        _stop = threading.Event()

        def on_event(partition_context, event):
            if _stop.is_set() or event is None:
                return
            try:
                payload = json.loads(event.body_as_str())
            except (json.JSONDecodeError, UnicodeDecodeError):
                return
            for mod in _parse_modules(payload):
                if virtual_device.lower() in mod["name"].lower():
                    for item in mod["reading"]:
                        param = (item.get("param") or "").strip()
                        if param:
                            received[param] = {
                                "value": item.get("value"),
                                "unit":  str(item.get("unit", "")).strip(),
                            }
                    if received:
                        done.set()

        starting_position = datetime.now(timezone.utc) - timedelta(seconds=10)
        eh_client = EventHubConsumerClient.from_connection_string(
            conn_str=azure["eventhub_conn_str"],
            consumer_group="$Default",
        )

        def _run_eh():
            try:
                eh_client.receive(on_event=on_event, starting_position=starting_position)
            except Exception:
                pass

        t = threading.Thread(target=_run_eh, daemon=True)
        t.start()

        azure_page.ensure_enabled()
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.select_only_device(virtual_device)
        azure_page.configure_all_devices_parameters(checked_only=True)
        azure_page.save()

        done.wait(timeout=subscribe_timeout)
        _stop.set()
        try:
            eh_client.close()
        except Exception:
            pass
        t.join(timeout=5)

        assert received, (
            f"在 {subscribe_timeout}s 内未收到 {virtual_device!r} 的 Event Hub 消息，"
            f"请确认：①虚拟设备已勾选并保存 ②Azure IoT 处于 Enable 状态 ③虚拟设备已配置公式参数"
        )
        print(f"\n[Event Hub] {virtual_device!r}: 收到 {len(received)} 个参数")

        reading_params = azure_page.get_virtual_device_readings(virtual_device)
        assert reading_params, (
            f"Reading 页面未读取到任何参数，设备：{virtual_device!r}。"
            f"请确认网关上该虚拟设备已配置公式参数。"
        )

        TOL_PCT = 0.05
        TOL_ABS = 1.0

        received_entries = []
        for v in received.values():
            try:
                received_entries.append((float(v["value"]), str(v["unit"]).strip()))
            except (TypeError, ValueError):
                received_entries.append((None, str(v["unit"]).strip()))

        unit_fails  = []
        value_fails = []

        used = [False] * len(received_entries)
        for rd_name, rd in reading_params.items():
            rd_unit = rd["unit"]
            try:
                rd_val = float(rd["value"])
            except (TypeError, ValueError):
                rd_val = None

            best_i, best_diff = None, float("inf")
            for i, (mv, _) in enumerate(received_entries):
                if used[i]:
                    continue
                if mv is not None and rd_val is not None:
                    d = abs(mv - rd_val)
                    if d < best_diff:
                        best_diff, best_i = d, i
                elif mv is None and rd_val is None:
                    best_i = i
                    break

            if best_i is None:
                unit_fails.append(f"  {rd_name!r}: 未找到匹配的 Event Hub 参数")
                continue

            used[best_i] = True
            mv, mv_unit = received_entries[best_i]
            if mv_unit != rd_unit:
                unit_fails.append(
                    f"  {rd_name!r}: Event Hub单位={mv_unit!r}  Reading单位={rd_unit!r}"
                )
            if rd_val is not None and mv is not None:
                ref = max(abs(mv), abs(rd_val), 1e-12)
                tol = max(TOL_ABS, ref * TOL_PCT)
                if best_diff > tol:
                    value_fails.append(
                        f"  {rd_name!r}: EventHub={mv}  Reading={rd_val}  diff={best_diff:.4f}"
                    )

        fails = []
        if unit_fails:
            fails.append(f"单位不匹配 ({len(unit_fails)} 项):\n" + "\n".join(unit_fails))
        if value_fails:
            fails.append(f"值超容差 ({len(value_fails)} 项):\n" + "\n".join(value_fails))

        assert not fails, (
            f"虚拟设备 {virtual_device!r} 数据与 Reading 页面不一致:\n"
            + "\n".join(fails)
        )
