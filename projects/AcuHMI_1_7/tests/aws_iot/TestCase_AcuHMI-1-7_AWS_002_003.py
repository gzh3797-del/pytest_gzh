"""
用例编号: TestCase_AcuHMI-1-7_AWS_002_003
用例标题: Interval 参数校验 — 各档 Interval 上报间隔验证

预置条件:
  1. 设备已正常上电，相关服务正常启动
  2. AWS IoT 连接参数及证书已通过 conftest aws_session_page 初始化

测试步骤（每档 Disable → 改 Interval → Enable → Save）:
  1. Disable AWS IoT 并保存，使设备停止推送
  2. Enable AWS IoT，将 Interval 修改为当前档位值
  3. 先建立 MQTT 订阅，记录时间基准 save_time
  4. Save，触发 AWS IoT 重新启动并以新 Interval 推送
  5. 轮询：只保留 device_timestamp >= save_time 的消息，收集 >= 3 条后退出
     超时上限 = 60s（重连预算）+ interval_secs × 3
  6. 验证所有相邻消息间隔（含第一个）严格等于配置值（±1s）

预期结果:
  - 每档 Interval 下，实际上报间隔（含切档后的首个间隔）必须等于配置值（±1s）
  - 整个用例完成后才执行 Disable（由 aws_session_page teardown 统一处理）
"""

import json
import logging
import threading
import time
from pathlib import Path

import pytest
import paho.mqtt.client as mqtt

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent      # AcuHMI-1-7/

# (UI 显示文本, 秒数) — 验证代表性档位
INTERVALS = [
    ("1 seconds",   1),
    ("10 seconds",  10),
    ("30 seconds",  30),
    ("60 seconds",  60),
]

_RECONNECT_BUDGET_SECS = 60


def _make_subscriber(cfg: dict):
    aws = cfg["aws_iot"]
    messages = []
    connected = threading.Event()

    def on_connect(client, userdata, flags, rc):
        if rc == 0:
            client.subscribe(aws["topic"], qos=1)
        connected.set()

    def on_message(client, userdata, msg):
        recv_time = time.time()
        device_ts = None
        try:
            payload = json.loads(msg.payload.decode("utf-8"))
            ts = payload.get("timestamp")
            if ts is not None:
                device_ts = float(ts)
        except (json.JSONDecodeError, UnicodeDecodeError, TypeError, ValueError):
            pass
        messages.append((device_ts, recv_time))

    client_id = f"pytest-sub-{int(time.time())}"
    try:
        client = mqtt.Client(
            client_id=client_id,
            transport="tcp",
            callback_api_version=mqtt.CallbackAPIVersion.VERSION1,
        )
    except (TypeError, AttributeError):
        client = mqtt.Client(client_id=client_id, transport="tcp")

    client.tls_set(
        ca_certs=str(_PROJECT_ROOT / aws["ca_file"]),
        certfile=str(_PROJECT_ROOT / aws["sub_cert_file"]),
        keyfile=str(_PROJECT_ROOT / aws["sub_key_file"]),
    )
    client.on_connect = on_connect
    client.on_message = on_message

    client.connect(aws["url"], port=8883, keepalive=60)
    client.loop_start()
    connected.wait(timeout=30)

    return client, messages


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


class TestCase_AcuHMI_1_7_AWS_002_003:

    @pytest.mark.aws_iot
    def test_all_interval_timings(self, aws_session_page, aws_cfg):
        """遍历所有 Interval 档位，每档 Disable → Enable + 改 Interval → Save，确保首条间隔准确。"""
        failures = []

        for interval_text, interval_secs in INTERVALS:
            logging.info(f">>> 开始验证 Interval={interval_text}")

            # Disable → Enable + 改 Interval，保证切档后计时器从零重新计时
            aws_session_page.disable()
            aws_session_page.ensure_enabled()
            aws_session_page.set_interval(interval_text)

            client, messages = _make_subscriber(aws_cfg)

            aws_session_page.save()
            save_time = time.time()

            deadline = save_time + _RECONNECT_BUDGET_SECS + interval_secs * 3
            poll_interval = min(1.0, interval_secs / 2)

            # 浏览器保活：每 60s ping 一次页面
            last_keepalive = save_time
            browser_alive = True
            while time.time() < deadline:
                if len(_get_fresh_timestamps(messages, save_time, interval_secs)) >= 3:
                    break
                time.sleep(poll_interval)
                if time.time() - last_keepalive >= 60.0:
                    try:
                        aws_session_page.page.evaluate("() => document.readyState")
                        last_keepalive = time.time()
                    except Exception as _ka_err:
                        logging.warning("浏览器保活 ping 失败：%s", _ka_err)
                        browser_alive = False
                        break

            client.loop_stop()
            client.disconnect()

            if not browser_alive:
                logging.warning("浏览器已关闭，仅验证已收集的 MQTT 数据")

            fresh_ts = _get_fresh_timestamps(messages, save_time, interval_secs)

            if len(fresh_ts) < 2:
                failures.append(
                    f"Interval={interval_text}：超时内仅收到 {len(fresh_ts)} 条新鲜消息，无法验证间隔"
                )
                continue

            gaps = [fresh_ts[i + 1] - fresh_ts[i] for i in range(len(fresh_ts) - 1)]
            # D/E 后计时器已重置，首个间隔同样需要符合配置值，无需跳过
            bad_gaps = [
                (idx + 1, g)
                for idx, g in enumerate(gaps)
                if abs(g - interval_secs) > 1
            ]
            if bad_gaps:
                failures.append(
                    f"Interval={interval_text}（{interval_secs}s），以下间隔不符（±1s）："
                    + ", ".join(f"第{i}个={g:.1f}s" for i, g in bad_gaps)
                )
            else:
                logging.info(f"<<< Interval={interval_text} 验证通过，间隔：{[f'{g:.1f}s' for g in gaps]}")

        # Disable 由 aws_session_page teardown 统一执行
        assert not failures, "以下 Interval 档位验证失败：\n" + "\n".join(failures)
