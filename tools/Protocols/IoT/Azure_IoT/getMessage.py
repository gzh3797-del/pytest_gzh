#!/usr/bin/env python3
"""
完整示例：监听 Azure IoT Hub 上所有设备的 Device-to-Cloud 消息
- 支持 Event Hub Consumer Group
- 输出消息内容和属性
- 可以重复消费，只要消息没过期
"""

import logging
from azure.eventhub import EventHubConsumerClient

# ==========================
# 配置区
# ==========================
# 1. IoT Hub 名称
IOT_HUB_NAME = "AccuenergyTest"

# 3. Consumer Group，默认 $Default，如果有多个服务建议独立创建
CONSUMER_GROUP = "$Default"

# 4. Event Hub 兼容连接字符串
# 在 Azure Portal -> IoT Hub -> Built-in Endpoints -> Events -> "Event Hub-compatible endpoint"
CONNECTION_STRING = "Endpoint=sb://ihsuprodblres064dednamespace.servicebus.windows.net/;SharedAccessKeyName=service;SharedAccessKey=kBE6EFMY+p5zjl9cexdnksLJ1+Lgr7tbAAIoTAbVFts=;EntityPath=iothub-ehub-accuenergy-70953761-f2919f1160"

# ==========================
# 日志配置
# ==========================
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# ==========================
# 消息处理回调
# ==========================
def on_event(partition_context, event):
    logging.info(f"收到设备消息，Partition:{partition_context.partition_id}")
    logging.info(f"消息内容: {event.body_as_str(encoding='UTF-8')}")
    if event.properties:
        logging.info(f"消息属性: {event.properties}")
    # 更新 checkpoint，防止重复处理
    partition_context.update_checkpoint(event)

# ==========================
# 主程序
# ==========================
def main():
    client = EventHubConsumerClient.from_connection_string(
        conn_str=CONNECTION_STRING,
        consumer_group=CONSUMER_GROUP,
    )

    logging.info("开始监听 IoT Hub 所有设备上传的消息...")
    try:
        with client:
            client.receive(
                on_event=on_event,
                starting_position="-1",  # 从最早的消息开始
            )
    except KeyboardInterrupt:
        logging.info("用户中断，停止监听")
    except Exception as e:
        logging.error(f"监听异常: {e}")

if __name__ == "__main__":
    main()