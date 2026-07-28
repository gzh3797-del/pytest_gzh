#!/bin/bash

# 订阅IoT消息
ENDPOINT="apbmxreennh9b-ats.iot.cn-northwest-1.amazonaws.com.cn"
CLIENT_ID="subscriber-$(date +%s)"
TOPIC="test/topic"

# 证书文件
CERT_FILE="29e287f92d07dc357838878febdac0cfc1d046b2abf81c715d2f42be4a2477c6-certificate.pem.crt"
KEY_FILE="29e287f92d07dc357838878febdac0cfc1d046b2abf81c715d2f42be4a2477c6-private.pem.key"
CA_FILE="AmazonRootCA1.pem"

echo "Subscribing to topic: $TOPIC"
echo "Press Ctrl+C to stop..."

mosquitto_sub \
  --host $ENDPOINT \
  --port 8883 \
  --id $CLIENT_ID \
  --topic $TOPIC \
  --cert $CERT_FILE \
  --key $KEY_FILE \
  --cafile $CA_FILE \
  --tls-version tlsv1.2