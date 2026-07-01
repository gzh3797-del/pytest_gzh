#!/usr/bin/env python3
"""
手动订阅 Azure IoT Hub 消息（通过 Event Hub 兼容端点）
用法：
    python tests/protocols/azure_iot/subscribe_messages.py
    python tests/protocols/azure_iot/subscribe_messages.py --timeout 120
    python tests/protocols/azure_iot/subscribe_messages.py --pretty

依赖：
    pip install azure-eventhub

配置：
    编辑 tests/protocols/azure_iot/config.yaml，填写 azure_iot.eventhub_conn_str
"""

import argparse
import json
import signal
import sys
import threading
from datetime import datetime, timezone, timedelta
from pathlib import Path

import yaml
from azure.eventhub import EventHubConsumerClient

_SCRIPT_DIR  = Path(__file__).resolve().parent
_CONFIG_PATH = _SCRIPT_DIR / "config.yaml"


def _load_cfg() -> dict:
    if not _CONFIG_PATH.exists():
        print(f"[ERROR] 找不到配置文件：{_CONFIG_PATH}", file=sys.stderr)
        sys.exit(1)
    return yaml.safe_load(_CONFIG_PATH.read_text(encoding="utf-8"))


def _fmt(payload: dict, pretty: bool) -> str:
    if pretty:
        return json.dumps(payload, ensure_ascii=False, indent=2)
    return json.dumps(payload, ensure_ascii=False)


def main():
    parser = argparse.ArgumentParser(description="手动订阅 Azure IoT Hub Event Hub 消息")
    parser.add_argument("--timeout", type=int, default=0,
                        help="监听超时秒数，0 表示持续监听直到 Ctrl+C（默认）")
    parser.add_argument("--pretty", action="store_true",
                        help="JSON 格式化输出")
    args = parser.parse_args()

    cfg   = _load_cfg()
    azure = cfg.get("azure_iot", {})
    conn_str = azure.get("eventhub_conn_str", "")
    if not conn_str or conn_str.startswith("Endpoint=sb://<"):
        print("[ERROR] 请先在 config.yaml 中填写合法的 azure_iot.eventhub_conn_str", file=sys.stderr)
        sys.exit(1)

    _stop = threading.Event()

    def on_event(partition_context, event):
        if _stop.is_set() or event is None:
            return
        now = datetime.now().strftime("%H:%M:%S")
        try:
            payload = json.loads(event.body_as_str())
            print(f"[{now}] {_fmt(payload, args.pretty)}")
        except (json.JSONDecodeError, UnicodeDecodeError):
            raw = event.body_as_str()
            print(f"[{now}] (raw) {raw[:200]}")

    starting_position = datetime.now(timezone.utc) - timedelta(seconds=10)
    client = EventHubConsumerClient.from_connection_string(
        conn_str=conn_str,
        consumer_group="$Default",
    )

    def _run():
        try:
            client.receive(on_event=on_event, starting_position=starting_position)
        except Exception:
            pass

    def _handle_sigint(sig, frame):
        print("\n[INFO] 收到 Ctrl+C，正在停止...", file=sys.stderr)
        _stop.set()

    signal.signal(signal.SIGINT, _handle_sigint)

    print(f"[INFO] 开始监听 Azure IoT Hub Event Hub 消息")
    print(f"[INFO] Consumer Group: $Default")
    print(f"[INFO] Starting Position: {starting_position.strftime('%Y-%m-%dT%H:%M:%SZ')}")
    if args.timeout > 0:
        print(f"[INFO] 超时：{args.timeout}s")
    print("[INFO] 按 Ctrl+C 停止...\n")

    t = threading.Thread(target=_run, daemon=True)
    t.start()

    if args.timeout > 0:
        t.join(timeout=args.timeout)
        _stop.set()
    else:
        _stop.wait()

    try:
        client.close()
    except Exception:
        pass

    print("\n[INFO] 已停止监听。")


if __name__ == "__main__":
    main()
