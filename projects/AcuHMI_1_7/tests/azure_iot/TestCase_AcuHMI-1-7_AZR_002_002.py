"""
用例编号: TestCase_AcuHMI-1-7_AZR_002_002
用例标题: Interval 参数校验 — 各档 Interval 上报间隔验证

预置条件:
  1. 设备已正常上电，相关服务正常启动
  2. Azure IoT 连接参数已通过 conftest azure_session_page 初始化

测试步骤（对每个 Interval 值执行一遍，全程共用同一浏览器）:
  1. 导航回 Azure IoT 页，Enable 并修改 Interval 为当前值
  2. 先建立 Event Hub 订阅，再保存配置（save() 完成时记录时间基准 save_time）
  3. 循环轮询：只保留 device_timestamp >= save_time 的消息
     当收集到 >= 3 条新鲜消息时提前退出；超时上限 = 60s + interval_secs × 3
  4. 验证相邻消息间隔严格等于配置值（±1s）
  5. 采集完成后立即 Disable，防止残留数据污染下一档

预期结果:
  - 每档 Interval 下，实际上报间隔必须等于配置值（±1s）
"""

import json
import threading
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path

import pytest
from azure.eventhub import EventHubConsumerClient

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent

INTERVALS = [
    ("1 seconds",   1),
    ("10 seconds",  10),
    ("30 seconds",  30),
    ("60 seconds",  60),
    ("90 seconds",  90),
    ("120 seconds", 120),
    ("180 seconds", 180),
    ("240 seconds", 240),
    ("300 seconds", 300),
    ("480 seconds", 480),
    ("600 seconds", 600),
]

_RECONNECT_BUDGET_SECS = 60


def _make_subscriber(cfg: dict):
    azure = cfg["azure_iot"]
    messages = []
    _stop = threading.Event()

    def on_event(partition_context, event):
        if _stop.is_set() or event is None:
            return
        recv_time = time.time()
        device_ts = None
        try:
            payload = json.loads(event.body_as_str())
            ts = payload.get("timestamp")
            if ts is not None:
                device_ts = float(ts)
        except (json.JSONDecodeError, UnicodeDecodeError, TypeError, ValueError):
            pass
        messages.append((device_ts, recv_time))

    starting_position = datetime.now(timezone.utc) - timedelta(seconds=10)
    client = EventHubConsumerClient.from_connection_string(
        conn_str=azure["eventhub_conn_str"],
        consumer_group="$Default",
    )

    def _run():
        try:
            client.receive(on_event=on_event, starting_position=starting_position)
        except Exception:
            pass

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return client, _stop, t, messages


def _get_fresh_timestamps(messages: list, save_time: float, interval_secs: float) -> list:
    seen: set = set()
    result = []
    for device_ts, recv_time in sorted(messages, key=lambda m: m[1]):
        if device_ts is not None:
            if device_ts < save_time:
                continue
            key = device_ts
            if key not in seen:
                seen.add(key)
                result.append(device_ts)
        else:
            if recv_time < save_time + 10:
                continue
            key = int(recv_time)
            if key not in seen:
                seen.add(key)
                result.append(recv_time)
    return sorted(result)


class TestCase_AcuHMI_1_7_AZR_002_002:

    @pytest.mark.azure_iot
    @pytest.mark.parametrize("interval_text,interval_secs", INTERVALS)
    def test_interval_timing(self, azure_session_page, azure_cfg, interval_text, interval_secs):
        azure_session_page.ensure_enabled()
        azure_session_page.set_interval(interval_text)

        client, _stop, t, messages = _make_subscriber(azure_cfg)

        azure_session_page.save()
        save_time = time.time()

        deadline = save_time + _RECONNECT_BUDGET_SECS + interval_secs * 3
        poll_interval = min(1.0, interval_secs / 2)
        while time.time() < deadline:
            if len(_get_fresh_timestamps(messages, save_time, interval_secs)) >= 3:
                break
            time.sleep(poll_interval)

        _stop.set()
        try:
            client.close()
        except Exception:
            pass
        t.join(timeout=5)

        azure_session_page.disable()

        fresh_ts = _get_fresh_timestamps(messages, save_time, interval_secs)

        assert len(fresh_ts) >= 2, (
            f"Interval={interval_text}：在超时上限（{_RECONNECT_BUDGET_SECS}s 重连预算"
            f" + {interval_secs * 3}s 采集窗口）内，仅收到 {len(fresh_ts)} 条"
            f" device_timestamp >= save_time 的新鲜消息，无法验证间隔"
        )

        gaps = [fresh_ts[i + 1] - fresh_ts[i] for i in range(len(fresh_ts) - 1)]
        fails = [
            (idx + 1, g)
            for idx, g in enumerate(gaps)
            if abs(g - interval_secs) > 1
        ]
        assert not fails, (
            f"Interval={interval_text}（{interval_secs}s），实际上报间隔必须等于配置值（±1s），"
            "以下间隔不符：" + ", ".join(f"第{i}个={g:.1f}s" for i, g in fails)
        )
