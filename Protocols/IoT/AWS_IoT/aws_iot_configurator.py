"""
AWS IoT 配置脚本
用法：
    python Protocols/IoT/AWS_IoT/aws_iot_configurator.py
    python Protocols/IoT/AWS_IoT/aws_iot_configurator.py --config Protocols/IoT/AWS_IoT/config.yaml
"""
import argparse
import logging
import os
import sys
import time

import yaml
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 项目根目录加入 sys.path
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
sys.path.insert(0, PROJECT_ROOT)

from test_case.WEB2_4100.operation.LoginPage import LoginPage
from test_case.WEB2_4100.operation.AWSIoTPage import AWSIoTPage

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')


def load_config(path: str) -> dict:
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def _abs_path(path: str) -> str:
    """相对路径转绝对路径（相对于项目根目录）"""
    if os.path.isabs(path):
        return path
    return os.path.abspath(os.path.join(PROJECT_ROOT, path))


def _create_driver() -> webdriver.Chrome:
    options = Options()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors=yes')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--disable-web-security')
    options.add_argument('--start-maximized')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    return webdriver.Chrome(options=options)


def run(config: dict):
    gw = config['gateway']
    aws = config['aws_iot']

    cert_file = _abs_path(aws['cert_file'])
    key_file = _abs_path(aws['key_file'])

    logging.info(f"Cert File: {cert_file}")
    logging.info(f"Key  File: {key_file}")

    if not os.path.exists(cert_file):
        raise FileNotFoundError(f"证书文件不存在：{cert_file}")
    if not os.path.exists(key_file):
        raise FileNotFoundError(f"私钥文件不存在：{key_file}")

    driver = _create_driver()
    try:
        driver.get(gw['url'])
        time.sleep(2)
        LoginPage(driver).login(gw['username'], gw['password'])
        # 等待 URL 从登录页跳转
        WebDriverWait(driver, 30).until(EC.url_changes(gw['url']))
        time.sleep(1)
        logging.info(f"登录成功，当前 URL: {driver.current_url}")
        time.sleep(2)

        page = AWSIoTPage(driver)
        page.navigate_to_aws_iot()
        result = page.configure(
            client_id=aws['client_id'],
            url=aws['url'],
            topic=aws['topic'],
            cert_file=cert_file,
            key_file=key_file,
            interval=aws.get('interval', '30 seconds'),
        )

        logging.info(f"Test Connection 页面反馈：{result}")
        if result:
            logging.info("✅ AWS IoT 配置并测试完成")
        time.sleep(3)
    finally:
        driver.quit()


def main():
    parser = argparse.ArgumentParser(description='AWS IoT 配置工具')
    parser.add_argument(
        '--config',
        default=os.path.join(os.path.dirname(__file__), 'config.yaml'),
        help='YAML 配置文件路径（默认：Protocols/IoT/AWS_IoT/config.yaml）'
    )
    args = parser.parse_args()
    config = load_config(args.config)
    run(config)


if __name__ == '__main__':
    main()
