import json
from pathlib import Path
from typing import List,Dict, Any
import pyautogui
import time
import os
import subprocess
import logging
import psutil
from typing import Optional, Tuple
import pytest
import cv2
import pyperclip
import pytesseract
import numpy as np


class AutoHelper:
    def __init__(self, failsafe=True, confidence=0.8):
        """
        初始化自动化助手

        Args:
            failsafe: 是否启用安全模式（鼠标到角落终止）
            confidence: 图像识别置信度
        """
        pyautogui.FAILSAFE = failsafe
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        self.confidence = confidence
        self.logger = logging.getLogger('AutoHelper')
        # 获取当前文件的绝对路径
        current_file = os.path.abspath(__file__)
        # 获取当前文件所在目录
        current_dir = os.path.dirname(current_file)
        # 获取上一级目录（去掉最后一级）
        self.project_root = os.path.dirname(current_dir)
        self.logger.info("🚀 自动化助手初始化完成")

    def launch_app(self, app_path: str, timeout: int = 30) -> bool:
        """
        启动应用程序（先进入目录再启动exe）

        Args:
            app_path: 应用程序完整路径
            timeout: 启动超时时间

        Returns:
            bool: 是否启动成功
        """
        try:
            # 获取应用目录和文件名
            app_dir = os.path.dirname(app_path)
            app_name = os.path.basename(app_path)

            self.logger.info(f"📱 启动应用: {app_name}")
            self.logger.info(f"📁 进入目录: {app_dir}")

            # 先进入目录，再启动应用
            process = subprocess.Popen(app_name, cwd=app_dir, shell=True)

            # 等待应用启动
            time.sleep(2)

            # 检查进程是否仍在运行
            if process.poll() is None:
                self.logger.info(f"✅ 应用启动成功: {app_name}")
                return True
            else:
                self.logger.error(f"❌ 应用启动失败，进程已退出: {app_name}")
                return False

        except Exception as e:
            self.logger.error(f"❌ 启动应用时出错: {e}")
            return False

    def kill_acuview_apps(self) -> bool:
        """
        关闭所有Acuview相关进程

        Returns:
            bool: 是否成功关闭所有进程
        """
        try:
            self.logger.info("🔪 开始关闭所有Acuview进程...")
            killed_count = 0

            # Acuview相关进程名
            acuview_process_names = [
                "Acuview 2.exe",
                "Acuview.exe",
                "Acuview2.exe",
                "Acuview"
            ]

            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    process_name = proc.info['name'].lower()
                    # 检查是否是Acuview相关进程
                    if any(acu_name.lower() in process_name for acu_name in acuview_process_names):
                        proc.terminate()
                        killed_count += 1
                        self.logger.info(f"✅ 已终止进程: {proc.info['name']} (PID: {proc.info['pid']})")

                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue

            # 等待进程完全结束
            time.sleep(2)

            # 强制杀死仍在运行的进程
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    process_name = proc.info['name'].lower()
                    if any(acu_name.lower() in process_name for acu_name in acuview_process_names):
                        proc.kill()
                        killed_count += 1
                        self.logger.warning(f"⚠️ 强制杀死进程: {proc.info['name']} (PID: {proc.info['pid']})")
                except:
                    continue

            self.logger.info(f"🎯 共关闭 {killed_count} 个Acuview进程")
            return killed_count > 0

        except Exception as e:
            self.logger.error(f"❌ 关闭Acuview进程时出错: {e}")
            return False

    def type_text(self, text: str, interval: float = 0.1):
        """
        输入文本

        Args:
            text: 要输入的文本
            interval: 每个字符输入的间隔时间
        """
        self.logger.info(f"⌨️ 输入文本: {text}")
        pyautogui.write(text, interval=interval)

    def hotkey(self, *keys):
        """
        执行快捷键

        Args:
            *keys: 快捷键组合，如 'ctrl', 'c'
        """
        key_str = '+'.join(keys)
        self.logger.info(f"🔧 执行快捷键: {key_str}")
        pyautogui.hotkey(*keys)

    def click_image(self, image_path: str, index: int = 0, offset_x: int = 0,
                    offset_y: int = 0, timeout: int = 3, confidence: float = 0.95) -> bool:
        """
        点击指定序号的相同图片元素

        Args:
            image_path: 图片路径
            index: 元素序号（0表示第一个，1表示第二个，以此类推）
            offset_x: X轴偏移
            offset_y: Y轴偏移
            timeout: 超时时间
            confidence: 置信度
        """
        if confidence is None:
            confidence = self.confidence

        full_image_path = self._get_full_image_path(image_path)

        self.logger.info(f"🔍 查找第 {index + 1} 个图片元素: {os.path.basename(full_image_path)}")

        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # 获取所有匹配的元素
                all_locations = list(pyautogui.locateAllOnScreen(full_image_path, confidence=confidence))

                if all_locations:
                    self.logger.info(f"📊 找到 {len(all_locations)} 个匹配元素")

                    if index < len(all_locations):
                        location = all_locations[index]
                        x, y = pyautogui.center(location)
                        center = (x + offset_x, y + offset_y)

                        self.logger.info(f"✅ 点击第 {index + 1} 个元素，位置: {center}")
                        pyautogui.click(center)
                        time.sleep(1)
                        return True
                    else:
                        pytest.fail(f"⚠️ 序号 {index} 超出范围，共找到 {len(all_locations)} 个元素")

            except:
                pytest.fail(f"❌ 未找到指定元素{image_path}（超时 {timeout} 秒)")

    def _get_full_image_path(self, image_path: str) -> str:
        """
        获取完整的图片路径

        Args:
            image_path: 图片路径

        Returns:
            str: 完整路径
        """
        if not os.path.isabs(image_path):
            # 如果是相对路径，转换为绝对路径
            full_image_path = os.path.join(self.project_root, image_path)
        else:
            # 如果是绝对路径，直接使用
            full_image_path = image_path

        # 确保有.png后缀
        if not full_image_path.lower().endswith('.png'):
            full_image_path += '.png'

        return full_image_path

    def get_mouse_position(self, duration: int = 0) -> Tuple[int, int]:
        """
        获取鼠标位置

        Args:
            duration: 监控持续时间（0表示立即返回）

        Returns:
            Tuple[int, int]: 鼠标坐标 (x, y)
        """
        if duration > 0:
            self.logger.info(f"🖱️ 监控鼠标位置 {duration} 秒，移动鼠标...")
            start_time = time.time()
            while time.time() - start_time < duration:
                x, y = pyautogui.position()
                print(f"坐标: ({x:4d}, {y:4d})", end='\r')
                time.sleep(0.1)
            print()  # 换行
        else:
            x, y = pyautogui.position()
            self.logger.info(f"🖱️ 当前鼠标位置: ({x}, {y})")

        return pyautogui.position()

    def wait(self, seconds: float):
        """
        等待指定时间

        Args:
            seconds: 等待秒数
        """
        self.logger.info(f"⏳ 等待 {seconds} 秒")
        time.sleep(seconds)

    def connect_device(self, device_image_path: str, timeout: int = 5, confidence: float = 0.95) -> bool:
        """
        通过两个参考元素定位并连接设备

        Args:
            device_image_path: 设备图片路径
            timeout: 超时时间（秒）
            confidence: 识别置信度

        Returns:
            bool: 是否成功连接设备
        """
        try:
            # 构建完整路径
            select_image_path = self._get_full_image_path('page_elements\Acuview_public\Add_Connect_page\Select.png')
            device_full_path = self._get_full_image_path(device_image_path)

            start_time = time.time()
            self.logger.info(f"🔗 开始连接设备流程，设备图片: {os.path.basename(device_full_path)}")

            while time.time() - start_time < timeout:
                try:
                    # 查找Select按钮
                    location1 = pyautogui.locateOnScreen(select_image_path, confidence=confidence)
                    if not location1:
                        self.logger.warning("⚠️ 未找到Select按钮，继续查找...")
                        time.sleep(1)
                        continue

                    # 查找设备图片
                    location2 = pyautogui.locateOnScreen(device_full_path, confidence=confidence)
                    if not location2:
                        self.logger.warning("⚠️ 未找到设备图片，继续查找...")
                        time.sleep(1)
                        continue

                    # 计算中心坐标
                    x1, y1 = pyautogui.center(location1)
                    x2, y2 = pyautogui.center(location2)

                    # 使用Select按钮的X坐标和设备图片的Y坐标
                    center = (x1, y2)

                    self.logger.info(f"📍 Select按钮位置: {location1}")
                    self.logger.info(f"📍 设备图片位置: {location2}")

                    # 安全检查点击位置
                    screen_width, screen_height = pyautogui.size()
                    if 0 <= center[0] <= screen_width and 0 <= center[1] <= screen_height:
                        pyautogui.click(center)
                        self.logger.info(f"🎯 点击位置: {center}")

                        # 等待界面响应
                        self.wait(1)

                        # 点击Connect按钮
                        if self.click_image('page_elements\Acuview_public\Add_Connect_page\Connect', timeout=5):
                            self.logger.info("✅ 设备连接成功")
                            self.wait(3)
                            return True
                        else:
                            self.logger.error("❌ 点击Connect按钮失败")
                    else:
                        self.logger.warning(f"🚫 点击位置超出屏幕范围: {center}")

                except Exception as e:
                    self.logger.warning(f"⚠️ 定位过程中出错: {e}")
                    time.sleep(1)

            self.logger.error(f"❌ 连接设备超时（{timeout}秒）")
            return False

        except Exception as e:
            self.logger.error(f"❌ 连接设备过程中发生异常: {e}")
            return False

    def is_process_running(self, process_name: str) -> bool:
        """
        检查进程是否在运行

        Args:
            process_name: 进程名

        Returns:
            bool: 是否在运行
        """
        try:
            for proc in psutil.process_iter(['name']):
                if process_name.lower() in proc.info['name'].lower():
                    return True
            return False
        except Exception as e:
            self.logger.warning(f"⚠️ 检查进程时出错: {e}")
            return False

    def check_image_exists(self, image_path: str, timeout: int = 2, confidence: float = 0.95) -> bool:
        """
        检查页面中是否存在指定图片

        Args:
            image_path: 图片路径
            timeout: 查找超时时间
            confidence: 置信度

        Returns:
            bool: 图片是否存在
        """
        if confidence is None:
            confidence = self.confidence

        full_image_path = self._get_full_image_path(image_path)

        self.logger.info(f"🔍 检查图片存在: {os.path.basename(full_image_path)}")

        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                location = pyautogui.locateOnScreen(full_image_path, confidence=confidence)
                if location:
                    self.logger.info(f"✅ 找到图片: {os.path.basename(full_image_path)}")
                    return True
            except Exception as e:
                pass
            time.sleep(1)

        self.logger.warning(f"⚠️ 未找到图片: {os.path.basename(full_image_path)}")
        return False

    def check_image_not_exists(self, image_path: str, timeout: int = 3, confidence: float = None) -> bool:
        """
        检查页面中是否不存在指定图片

        Args:
            image_path: 图片路径
            timeout: 检查超时时间
            confidence: 置信度

        Returns:
            bool: 图片是否不存在
        """
        if confidence is None:
            confidence = self.confidence

        full_image_path = self._get_full_image_path(image_path)

        self.logger.info(f"🔍 检查图片不存在: {os.path.basename(full_image_path)}")

        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                location = pyautogui.locateOnScreen(full_image_path, confidence=confidence)
                if not location:
                    self.logger.info(f"✅ 图片不存在: {os.path.basename(full_image_path)}")
                    return True
            except Exception as e:
                self.logger.warning(f"⚠️ 检查图片时出错: {e}")
            time.sleep(1)

        self.logger.warning(f"⚠️ 图片仍然存在: {os.path.basename(full_image_path)}")
        return False

    def paste_text(self, text: str, delay: float = 1):
        """
        将文本复制到剪贴板并粘贴

        Args:
            text: 要粘贴的文本
            delay: 粘贴前后的延迟（秒）
        """
        try:
            self.logger.info(f"📋 准备粘贴文本: {text}")

            # 将文本复制到剪贴板
            pyperclip.copy(text)

            # 短暂延迟确保复制完成
            time.sleep(delay)

            # 执行粘贴操作 (Ctrl+V)
            pyautogui.hotkey('ctrl', 'v')

            # 等待粘贴完成
            time.sleep(delay)

            self.logger.info("✅ 文本粘贴完成")

        except Exception as e:
            self.logger.error(f"❌ 文本粘贴失败: {e}")
            raise

    def click_pos(self, pos):
        pyautogui.click(pos)
        self.wait(1)
        self.logger.info(f'点击坐标{pos}')

    def double_click_pos(self, pos):
        pyautogui.click(pos)
        pyautogui.click(pos)
        self.wait(1)
        self.logger.info(f'双击坐标{pos}')

    def copy_echilog_info(self, time, type, old_value, new_value):
        """
        点击四个坐标点分别复制时间、类型、旧值、新值

        Args:
            pos_time: 时间坐标 (x, y)
            pos_type: 类型坐标 (x, y)
            pos_old_value: 旧值坐标 (x, y)
            pos_new_value: 新值坐标 (x, y)

        Returns:
            tuple: (time_value, type_value, old_value, new_value)
        """
        # 复制时间
        pyautogui.click(time)
        self.wait(0.5)
        pyautogui.hotkey('ctrl', 'c')
        self.wait(0.5)
        time_value = pyperclip.paste()

        # 复制类型
        pyautogui.click(type)
        self.wait(0.5)
        pyautogui.hotkey('ctrl', 'c')
        self.wait(0.5)
        type_value = pyperclip.paste()

        # 复制旧值
        pyautogui.click(old_value)
        self.wait(0.5)
        pyautogui.hotkey('ctrl', 'c')
        self.wait(0.5)
        old_value = pyperclip.paste()

        # 复制新值
        pyautogui.click(new_value)
        self.wait(0.5)
        pyautogui.hotkey('ctrl', 'c')
        self.wait(0.5)
        new_value = pyperclip.paste()
        return time_value, type_value, old_value, new_value

    def parse_ocmf(self, data):
        # 分割 OCMF 格式的数据
        parts = data.split('|', 2)  # 最多分割成3部分
        if len(parts) >= 2:
            # 第一部分是 OCMF 头，第二部分是 JSON 数据
            json_str = parts[1]
            # 解析 JSON
            data_dict = json.loads(json_str)
            return data_dict
        return None

    def extract_and_ocr_by_coordinates(self,
                                       top_left: Tuple[int, int],
                                       bottom_right: Tuple[int, int],
                                       description: str = "",
                                       lang: str = 'chi_sim+eng') -> dict:
        """
        提取指定区域图像并进行OCR文字识别（不保存图片）

        Args:
            top_left: 左上角坐标 (x, y)
            bottom_right: 右下角坐标 (x, y)
            description: 图像描述
            lang: OCR语言 (chi_sim: 简体中文, eng: 英文)

        Returns:
            dict: 包含识别结果的字典
        """
        try:
            # 验证坐标有效性
            if (top_left[0] >= bottom_right[0] or top_left[1] >= bottom_right[1]):
                self.logger.error("❌ 坐标无效：左上角坐标应小于右下角坐标")
                return {"success": False, "error": "坐标无效"}

            # 计算区域尺寸
            width = bottom_right[0] - top_left[0]
            height = bottom_right[1] - top_left[1]

            if width <= 0 or height <= 0:
                self.logger.error("❌ 区域尺寸无效")
                return {"success": False, "error": "区域尺寸无效"}

            # 截取指定区域
            screenshot = pyautogui.screenshot(region=(top_left[0], top_left[1], width, height))

            # 将PIL图像转换为OpenCV格式（提高OCR精度）
            img_array = np.array(screenshot)
            img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)

            # 图像预处理（提高OCR识别率）
            # 转换为灰度图
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            cv2.imwrite('gray_image.png', gray)

            # 可选：应用二值化
            # _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            # 进行OCR识别
            text = pytesseract.image_to_string(gray, lang=lang)

            # 清理文本
            cleaned_text = text.strip()

            result = {
                "success": True,
                "text": cleaned_text,
                "coordinates": {
                    "top_left": top_left,
                    "bottom_right": bottom_right,
                    "width": width,
                    "height": height
                },
                "description": description,
                "char_count": len(cleaned_text)
            }

            self.logger.info(f"🔤 OCR识别完成: 识别到 {len(cleaned_text)} 个字符")
            if cleaned_text:
                self.logger.info(f"📄 识别内容: {cleaned_text[:100]}{'...' if len(cleaned_text) > 100 else ''}")

            return result

        except Exception as e:
            self.logger.error(f"❌ OCR识别失败: {e}")
            return {"success": False, "error": str(e)}

    def extract_multiple_regions_text(self,
                                      regions: List[dict],
                                      lang: str = 'chi_sim+eng') -> List[dict]:
        """
        批量提取多个区域文字（不保存图片）

        Args:
            regions: 区域列表，每个区域包含:
                    - top_left: 左上角坐标
                    - bottom_right: 右下角坐标
                    - name: 区域名称（可选）
                    - description: 描述（可选）
            lang: OCR语言

        Returns:
            List[dict]: 识别结果列表
        """
        results = []

        for i, region in enumerate(regions):
            top_left = region.get('top_left')
            bottom_right = region.get('bottom_right')
            name = region.get('name', f"region_{i + 1}")
            description = region.get('description', f"区域 {i + 1}")

            if not top_left or not bottom_right:
                self.logger.warning(f"⚠️ 跳过区域 {i + 1}: 缺少坐标信息")
                results.append({
                    "region_index": i + 1,
                    "region_name": name,
                    "success": False,
                    "error": "缺少坐标信息"
                })
                continue

            result = self.extract_and_ocr_by_coordinates(
                top_left=top_left,
                bottom_right=bottom_right,
                description=description,
                lang=lang
            )

            # 添加区域信息到结果中
            result.update({
                "region_index": i + 1,
                "region_name": name
            })

            results.append(result)

        success_count = len([r for r in results if r.get('success', False)])
        self.logger.info(f"📦 批量OCR完成: 成功 {success_count}/{len(regions)} 个区域")
        return results

    def extract_image_by_coordinates(self,
                                     top_left: Tuple[int, int],
                                     bottom_right: Tuple[int, int],
                                     description: str = "") -> np.ndarray:
        """
        提取屏幕指定区域的图像（不保存文件，返回numpy数组）

        Args:
            top_left: 左上角坐标 (x, y)
            bottom_right: 右下角坐标 (x, y)
            description: 图像描述

        Returns:
            np.ndarray: 图像数组
        """
        try:
            # 验证坐标有效性
            if (top_left[0] >= bottom_right[0] or top_left[1] >= bottom_right[1]):
                self.logger.error("❌ 坐标无效：左上角坐标应小于右下角坐标")
                return None

            # 计算区域尺寸
            width = bottom_right[0] - top_left[0]
            height = bottom_right[1] - top_left[1]

            if width <= 0 or height <= 0:
                self.logger.error("❌ 区域尺寸无效")
                return None

            # 截取指定区域
            screenshot = pyautogui.screenshot(region=(top_left[0], top_left[1], width, height))
            img_array = np.array(screenshot)

            self.logger.info(f"🖼️  图像提取成功")
            self.logger.info(f"📍 区域: ({top_left[0]}, {top_left[1]}) 到 ({bottom_right[0]}, {bottom_right[1]})")
            self.logger.info(f"📏 尺寸: {width} x {height} 像素")
            self.logger.info(f"📝 描述: {description}")

            return img_array

        except Exception as e:
            self.logger.error(f"❌ 图像提取失败: {e}")
            return None

    def check_csv_file(self, file_path, expected_data, expected_data_len):
        # 读取CSV文件
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        with open(file_path, 'r', encoding='utf-8') as file:
            line_count = sum(1 for line in file) - 3
        third_line = lines[2].strip()
        if ',' in third_line:
            actual_ocmf_part = third_line.split(',', 1)[1][:-1]
        # 方法2：使用制表符分隔
        elif '\t' in third_line:
            actual_ocmf_part = third_line.split('\t', 1)[1][:-1]
        pytest.assume(actual_ocmf_part == expected_data, "上位机读取的第一条数据和csv文件第一条数据不匹配")
        pytest.assume(line_count == expected_data_len,
                      f"上位机读取的数据总数{expected_data_len}，csv文件数据总数{line_count}")

    def load_pos_config(self) -> Dict[str, Any]:
        """
        加载坐标配置文件

        Args:
            config_path: 配置文件路径，如果为None则使用默认路径

        Returns:
            Dict[str, Any]: 配置字典
        """

        # 默认配置文件路径
        config_path = Path(__file__).parent.parent / "config" / "pos_config.json"
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        self.logger.info(f"✅ 配置文件加载成功: {config_path}")
        self.logger.info(f"📊 共加载 {len(config)} 个配置项")
        return config

    def quick_ocr_by_config(self, config_name: str) -> str:
        """
        根据配置名称快速OCR识别

        Args:
            config_name: 配置项名称
            config_path: 配置文件路径

        Returns:
            str: 识别到的文本
        """
        try:
            config = self.load_pos_config()

            if not config:
                self.logger.error("❌ 配置加载失败")
                return ""

            if config_name not in config:
                self.logger.error(f"❌ 配置项 '{config_name}' 不存在")
                available_keys = list(config.keys())
                self.logger.info(f"📋 可用配置项: {available_keys}")
                return ""

            item_config = config[config_name]

            # 提取配置参数
            top_left = tuple(item_config.get('top_left', [0, 0]))
            bottom_right = tuple(item_config.get('bottom_right', [0, 0]))
            lang = item_config.get('lang', 'chi_sim+eng')
            description = item_config.get('description', config_name)

            self.logger.info(f"🔍 根据配置 '{config_name}' 进行OCR识别")
            self.logger.info(f"📍 区域: {top_left} -> {bottom_right}")
            self.logger.info(f"🌐 语言: {lang}")
            self.logger.info(f"📝 描述: {description}")

            # 调用quick_ocr方法
            result = self.quick_ocr_enhanced(
                top_left=top_left,
                bottom_right=bottom_right,
                lang=lang
            )

            return result

        except Exception as e:
            self.logger.error(f"❌ 根据配置OCR识别失败: {e}")
            return ""

    def batch_ocr_by_config(self, config_names: list) -> Dict[str, str]:
        """
        批量根据配置进行OCR识别

        Args:
            config_names: 配置项名称列表
            config_path: 配置文件路径

        Returns:
            Dict[str, str]: 识别结果字典 {配置名: 识别文本}
        """
        results = {}

        for config_name in config_names:
            text = self.quick_ocr_by_config(config_name)
            results[config_name] = text

        self.logger.info(f"📦 批量OCR完成: 成功 {len([v for v in results.values() if v])}/{len(config_names)} 个配置项")
        return results

    def get_all_config_names(self) -> list:
        """
        获取所有配置项名称

        Args:
            config_path: 配置文件路径

        Returns:
            list: 配置项名称列表
        """
        config = self.load_pos_config()
        return list(config.keys())

    def validate_config_item(self, config_name: str) -> bool:
        """
        验证配置项是否有效

        Args:
            config_name: 配置项名称
            config_path: 配置文件路径

        Returns:
            bool: 配置项是否有效
        """
        config = self.load_pos_config()

        if config_name not in config:
            return False

        item = config[config_name]
        required_fields = ['top_left', 'bottom_right']

        for field in required_fields:
            if field not in item:
                return False

        # 验证坐标有效性
        top_left = item['top_left']
        bottom_right = item['bottom_right']

        if (len(top_left) != 2 or len(bottom_right) != 2 or
                top_left[0] >= bottom_right[0] or top_left[1] >= bottom_right[1]):
            return False

        return True

    # 确保 quick_ocr 方法存在（如果还没有的话）
    def quick_ocr_enhanced(self,
                           top_left: Tuple[int, int],
                           bottom_right: Tuple[int, int],
                           lang: str = 'chi_sim+eng') -> str:
        """
        增强版OCR识别，调整引擎参数
        """
        try:
            # 截取区域
            width = bottom_right[0] - top_left[0]
            height = bottom_right[1] - top_left[1]
            screenshot = pyautogui.screenshot(region=(top_left[0], top_left[1], width, height))

            # 转换为OpenCV格式并进行图像处理
            img_array = np.array(screenshot)
            img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)

            # 多种图像预处理
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

            # 方法1: 二值化
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            # 方法2: 调整对比度
            alpha = 1.5  # 对比度控制
            beta = 0  # 亮度控制
            enhanced = cv2.convertScaleAbs(gray, alpha=alpha, beta=beta)
            # cv2.imwrite('gray_image.png', enhanced)
            # 尝试不同的图像预处理方法
            images_to_try = [
                ("original", gray),
                ("binary", binary),
                ("enhanced", enhanced)
            ]

            best_result = ""

            for method_name, processed_img in images_to_try:
                try:
                    # 配置OCR引擎参数
                    custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz '

                    text = pytesseract.image_to_string(processed_img, lang=lang, config=custom_config)
                    cleaned_text = text.strip()

                    if cleaned_text:
                        self.logger.debug(f"方法 {method_name} 识别结果: {cleaned_text}")
                        if not best_result or len(cleaned_text) > len(best_result):
                            best_result = cleaned_text

                except Exception as e:
                    self.logger.debug(f"方法 {method_name} 失败: {e}")
                    continue

            return best_result

        except Exception as e:
            self.logger.error(f"❌ 增强OCR失败: {e}")
            return ""

    def validate_json(self, input_data: dict) -> Dict[str, Any]:
        """
        检查JSON字符串中的字段是否存在，并使用pytest.assume进行断言

        Args:
            input_data: 输入的JSON字符串或字典

        Returns:
            Dict: 验证结果，包含断言结果和详细信息
        """
        # 预期的字段列表
        expected_fields = {
            "FV", "GI", "GS", "GV", "PG", "MV", "MM", "MS", "MF",
            "IS", "IL", "IF", "IT", "ID", "CT", "CI", "TT", "RD"
        }

        # 存储断言结果
        assertion_results = {
            "all_passed": True,
            "failed_assertions": [],
            "details": {}
        }

        try:
            # 如果输入是字符串，尝试解析为JSON
            if isinstance(input_data, str):
                data = json.loads(input_data)
            else:
                data = input_data

            # 获取实际存在的字段
            actual_fields = set(data.keys())

            # 找出缺失的字段
            missing_fields = expected_fields - actual_fields
            existing_fields = expected_fields & actual_fields

            # 断言1: 检查是否所有预期字段都存在
            try:
                pytest.assume(len(missing_fields) == 0,
                              f"缺失字段: {missing_fields}")
                assertion_results["details"]["missing_fields"] = list(missing_fields)
            except AssertionError as e:
                assertion_results["all_passed"] = False
                assertion_results["failed_assertions"].append(str(e))

            # 断言2: 检查RD字段是否存在且是列表
            try:
                pytest.assume("RD" in data, "RD字段缺失")
                if "RD" in data:
                    pytest.assume(isinstance(data["RD"], list), "RD字段不是列表类型")
                    assertion_results["details"]["rd_exists"] = True
                    assertion_results["details"]["rd_is_list"] = True
                else:
                    assertion_results["details"]["rd_exists"] = False
                    assertion_results["details"]["rd_is_list"] = False
            except AssertionError as e:
                assertion_results["all_passed"] = False
                assertion_results["failed_assertions"].append(str(e))

            # 验证RD字段的结构（如果存在）
            rd_validation = {}
            if "RD" in data and isinstance(data["RD"], list):
                rd_items = data["RD"]
                rd_validation = {
                    "count": len(rd_items),
                    "sample_structure": None
                }

                # 断言3: 检查RD数组不为空
                try:
                    pytest.assume(len(rd_items) > 0, "RD数组为空")
                    assertion_results["details"]["rd_count"] = len(rd_items)
                except AssertionError as e:
                    assertion_results["all_passed"] = False
                    assertion_results["failed_assertions"].append(str(e))

                if rd_items:
                    # 检查第一个RD项目的字段
                    rd_expected_fields = {"TM", "TX", "RV", "RI", "RU", "RT", "EF", "ST"}
                    first_rd = rd_items[0]

                    # 断言4: 检查RD项目是字典类型
                    try:
                        pytest.assume(isinstance(first_rd, dict), "RD项目不是字典类型")
                        assertion_results["details"]["rd_item_is_dict"] = True
                    except AssertionError as e:
                        assertion_results["all_passed"] = False
                        assertion_results["failed_assertions"].append(str(e))

                    if isinstance(first_rd, dict):
                        rd_actual_fields = set(first_rd.keys())
                        missing_rd_fields = rd_expected_fields - rd_actual_fields

                        # 断言5: 检查RD项目是否包含所有必需字段
                        try:
                            pytest.assume(len(missing_rd_fields) == 0,
                                          f"RD项目缺失字段: {missing_rd_fields}")
                            assertion_results["details"]["missing_rd_fields"] = list(missing_rd_fields)
                        except AssertionError as e:
                            assertion_results["all_passed"] = False
                            assertion_results["failed_assertions"].append(str(e))

                        # 特别检查TM和TX字段
                        try:
                            pytest.assume("TM" in first_rd, "RD项目缺失TM字段")
                            pytest.assume("TX" in first_rd, "RD项目缺失TX字段")
                            assertion_results["details"]["has_tm"] = "TM" in first_rd
                            assertion_results["details"]["has_tx"] = "TX" in first_rd
                        except AssertionError as e:
                            assertion_results["all_passed"] = False
                            assertion_results["failed_assertions"].append(str(e))

                        rd_validation["sample_structure"] = {
                            "expected": list(rd_expected_fields),
                            "actual": list(rd_actual_fields),
                            "missing": list(missing_rd_fields)
                        }

            # 汇总结果
            assertion_results["details"].update({
                "existing_fields": list(existing_fields),
                "total_expected": len(expected_fields),
                "total_found": len(existing_fields),
                "rd_validation": rd_validation,
                "data_sample": {field: data.get(field, "MISSING") for field in list(existing_fields)[:3]}
            })

            return assertion_results

        except json.JSONDecodeError as e:
            error_msg = f"JSON解析错误: {e}"
            pytest.assume(False, error_msg)
            return {
                "all_passed": False,
                "failed_assertions": [error_msg],
                "details": {
                    "error": error_msg,
                    "missing_fields": list(expected_fields),
                    "existing_fields": []
                }
            }
        except Exception as e:
            error_msg = f"验证过程中出错: {e}"
            pytest.assume(False, error_msg)
            return {
                "all_passed": False,
                "failed_assertions": [error_msg],
                "details": {
                    "error": error_msg,
                    "missing_fields": list(expected_fields),
                    "existing_fields": []
                }
            }


if __name__ == "__main__":
    # 初始化助手
    # 初始化助手
    helper = AutoHelper()
    # time.sleep(3)
    # # 快速OCR识别
    # r = helper.quick_ocr_by_config('Max Transaction Id')
    # print(r)
    helper.check_csv_file(r'C:\Users\YiSong\Acuview2\Export\DE55061235_Transaction Log_2025_11_24_10_10_36.csv', 111,
                          326)
