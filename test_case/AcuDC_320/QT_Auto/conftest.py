"""
pytest配置文件 - 集成了截图功能、Excel报告生成和Windows防睡眠
专门针对AutoHelper和QT应用的截图功能
"""

import pytest
import sys
import os
import pandas as pd
import ctypes
import json
from datetime import datetime
from pathlib import Path
import traceback
import pyautogui

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Windows防睡眠类
class WindowsKeepAwake:
    """Windows系统防睡眠工具"""

    def __init__(self):
        self.is_windows = sys.platform == "win32"
        if self.is_windows:
            self.ES_CONTINUOUS = 0x80000000
            self.ES_SYSTEM_REQUIRED = 0x00000001
            self.ES_DISPLAY_REQUIRED = 0x00000002

    def __enter__(self):
        if self.is_windows:
            ctypes.windll.kernel32.SetThreadExecutionState(
                self.ES_CONTINUOUS | self.ES_SYSTEM_REQUIRED | self.ES_DISPLAY_REQUIRED
            )
            print("🔋 防睡眠已启用")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.is_windows:
            ctypes.windll.kernel32.SetThreadExecutionState(self.ES_CONTINUOUS)
            print("🔋 防睡眠已禁用")

# 测试结果收集器
class TestResultCollector:
    """收集测试结果并生成多种格式报告"""

    def __init__(self):
        self.test_results = []
        self.test_stats = {
            "total": 0,
            "passed": 0,
            "failed": 0,
            "skipped": 0,
            "error": 0,
            "duration": 0.0
        }

    def add_result(self, report, screenshot_path=None):
        """添加测试结果"""
        if report.when == "call":
            test_full_name = report.nodeid

            # 解析测试文件路径和测试函数名称
            if "::" in test_full_name:
                test_file_part = test_full_name.split("::")[0]
                test_class_func = test_full_name.split("::")[1:]
                test_name = "::".join(test_class_func)

                # 提取文件名（不含路径和扩展名）
                test_file_name = os.path.splitext(os.path.basename(test_file_part))[0]
            else:
                test_file_name = "unknown_test_file"
                test_name = test_full_name

            if test_name.startswith("test_"):
                display_name = test_name[5:]
            else:
                display_name = test_name

            # 确定测试状态
            if report.passed:
                status = "通过"
                self.test_stats["passed"] += 1
            elif report.failed:
                status = "失败"
                self.test_stats["failed"] += 1
            elif report.skipped:
                status = "跳过"
                self.test_stats["skipped"] += 1
            else:
                status = "错误"
                self.test_stats["error"] += 1

            # 构建结果字典
            result_dict = {
                "测试文件": test_file_name,
                "用例名称": display_name,
                "是否通过": status,
                "执行时间(秒)": round(report.duration, 3),
                "错误信息": str(report.longrepr) if report.failed else ""
            }

            # 如果有截图路径，添加到结果中
            if screenshot_path:
                # 使用绝对路径
                result_dict["截图路径"] = str(screenshot_path.resolve())

            self.test_results.append(result_dict)
            self.test_stats["total"] += 1
            self.test_stats["duration"] += report.duration

    def generate_excel_report(self, report_dir):
        """生成Excel报告"""
        if not self.test_results:
            return None

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_file = report_dir / f"test_report_{timestamp}.xlsx"

        df = pd.DataFrame(self.test_results)

        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            # 测试报告详情表
            df.to_excel(writer, sheet_name='测试详情', index=False)

            # 测试统计表
            summary_data = {
                "统计项目": ["总用例数", "通过数", "失败数", "跳过数", "错误数", "执行时间(秒)", "通过率"],
                "数值": [
                    self.test_stats["total"],
                    self.test_stats["passed"],
                    self.test_stats["failed"],
                    self.test_stats["skipped"],
                    self.test_stats["error"],
                    round(self.test_stats["duration"], 3),
                    f"{self.test_stats['passed']/self.test_stats['total']*100:.2f}%" if self.test_stats['total'] > 0 else "0%"
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='测试统计', index=False)

            # 测试文件统计表
            file_stats = {}
            for result in self.test_results:
                test_file = result["测试文件"]
                status = result["是否通过"]

                if test_file not in file_stats:
                    file_stats[test_file] = {"total": 0, "passed": 0, "failed": 0, "skipped": 0, "error": 0}

                file_stats[test_file]["total"] += 1
                if status == "通过":
                    file_stats[test_file]["passed"] += 1
                elif status == "失败":
                    file_stats[test_file]["failed"] += 1
                elif status == "跳过":
                    file_stats[test_file]["skipped"] += 1
                elif status == "错误":
                    file_stats[test_file]["error"] += 1

            file_data = []
            for file_name, stats in file_stats.items():
                pass_rate = stats["passed"] / stats["total"] * 100 if stats["total"] > 0 else 0
                file_data.append({
                    "测试文件": file_name,
                    "总用例数": stats["total"],
                    "通过数": stats["passed"],
                    "失败数": stats["failed"],
                    "跳过数": stats["skipped"],
                    "错误数": stats["error"],
                    "通过率": f"{pass_rate:.2f}%"
                })

            if file_data:
                file_df = pd.DataFrame(file_data)
                file_df.to_excel(writer, sheet_name='文件统计', index=False)

            # 设置列宽
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                if sheet_name == '测试详情':
                    worksheet.column_dimensions['A'].width = 25
                    worksheet.column_dimensions['B'].width = 30
                    worksheet.column_dimensions['C'].width = 15
                    worksheet.column_dimensions['D'].width = 15
                    worksheet.column_dimensions['E'].width = 50
                    if "截图路径" in df.columns:
                        worksheet.column_dimensions['F'].width = 100  # 增加宽度以显示完整路径
                elif sheet_name == '测试统计':
                    worksheet.column_dimensions['A'].width = 20
                    worksheet.column_dimensions['B'].width = 20
                elif sheet_name == '文件统计':
                    worksheet.column_dimensions['A'].width = 30
                    worksheet.column_dimensions['B'].width = 15
                    worksheet.column_dimensions['C'].width = 15
                    worksheet.column_dimensions['D'].width = 15
                    worksheet.column_dimensions['E'].width = 15
                    worksheet.column_dimensions['F'].width = 15
                    worksheet.column_dimensions['G'].width = 15

        print(f"📊 Excel报告已生成: {excel_file}")
        return excel_file

    def generate_json_report(self, report_dir):
        """生成JSON报告"""
        if not self.test_results:
            return None

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_file = report_dir / f"test_report_{timestamp}.json"

        report_data = {
            "timestamp": datetime.now().isoformat(),
            "statistics": self.test_stats,
            "test_results": self.test_results
        }

        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, ensure_ascii=False, indent=2)

        print(f"📋 JSON报告已生成: {json_file}")
        return json_file

    def print_summary(self, report_dir):
        """打印测试摘要"""
        print("\n" + "="*50)
        print("📊 测试摘要")
        print("="*50)
        print(f"📁 报告目录: {report_dir}")
        print(f"📈 总计用例: {self.test_stats['total']}")
        print(f"✅ 通过: {self.test_stats['passed']}")
        print(f"❌ 失败: {self.test_stats['failed']}")
        print(f"⏭️  跳过: {self.test_stats['skipped']}")
        print(f"⚠️  错误: {self.test_stats['error']}")

        if self.test_stats['total'] > 0:
            pass_rate = self.test_stats['passed'] / self.test_stats['total'] * 100
            print(f"📊 通过率: {pass_rate:.2f}%")
            print(f"⏱️  总耗时: {self.test_stats['duration']:.3f}秒")
        print("="*50)

# 截图管理器
class ScreenshotManager:
    """管理测试截图保存 - 使用pyautogui进行截图"""

    def __init__(self, base_report_dir):
        self.base_report_dir = base_report_dir
        self.screenshot_counter = {}
        # 创建screenshot根目录
        self.screenshot_root = base_report_dir / "screenshots"
        self.screenshot_root.mkdir(parents=True, exist_ok=True)
        print(f"📸 截图目录: {self.screenshot_root}")

    def _sanitize_filename(self, filename):
        """清理文件名中的非法字符"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        # 移除多余的空格和特殊字符
        filename = filename.strip().replace(' ', '_')
        # 限制文件名长度
        if len(filename) > 100:
            filename = filename[:100]
        return filename

    def get_screenshot_path(self, test_full_name, is_failed=False):
        """获取截图保存路径，按照 测试文件/测试函数 层级组织"""
        # 解析测试完整名称
        if "::" in test_full_name:
            # 格式: test_file.py::TestClass::test_function
            parts = test_full_name.split("::")
            test_file_path = parts[0]

            # 提取文件名（不含路径和扩展名）
            test_file_name = os.path.splitext(os.path.basename(test_file_path))[0]

            # 如果有类名和函数名
            if len(parts) >= 2:
                # 处理测试函数名（可能包含类名）
                test_function = parts[-1]

                # 如果包含类名
                if len(parts) >= 3:
                    test_class = parts[1]
                    # 清理类名中的非法字符
                    test_class = self._sanitize_filename(test_class)
                    # 创建测试文件目录下的类名目录
                    function_dir_name = f"{test_class}_{test_function}"
                else:
                    function_dir_name = test_function
            else:
                function_dir_name = "unnamed_test"
        else:
            # 没有::分隔符的情况
            test_file_name = "unknown_test_file"
            function_dir_name = test_full_name

        # 清理文件名
        test_file_name = self._sanitize_filename(test_file_name)
        function_dir_name = self._sanitize_filename(function_dir_name)

        # 创建目录结构
        test_file_dir = self.screenshot_root / test_file_name
        test_function_dir = test_file_dir / function_dir_name

        # 确保目录存在
        test_function_dir.mkdir(parents=True, exist_ok=True)

        # 生成唯一的截图文件名
        key = f"{test_file_name}::{function_dir_name}"
        if key not in self.screenshot_counter:
            self.screenshot_counter[key] = 1
        else:
            self.screenshot_counter[key] += 1

        counter = self.screenshot_counter[key]

        # 根据测试状态命名文件
        status = "FAILED" if is_failed else "PASSED"
        timestamp = datetime.now().strftime("%H%M%S_%f")[:-3]  # 毫秒级时间戳

        # 生成更易读的文件名
        screenshot_name = f"{status}_{timestamp}_{counter:03d}.png"

        screenshot_file = test_function_dir / screenshot_name
        return screenshot_file

    def take_screenshot(self, test_full_name, is_failed=False, description=""):
        """使用pyautogui截图"""
        try:
            # 获取保存路径
            screenshot_path = self.get_screenshot_path(test_full_name, is_failed)

            print(f"📸 准备截图: {description}")
            print(f"📸 截图路径: {screenshot_path}")

            # 使用pyautogui截图
            screenshot = pyautogui.screenshot()

            # 保存截图
            screenshot.save(str(screenshot_path))

            # 验证文件是否保存成功
            if screenshot_path.exists():
                file_size = screenshot_path.stat().st_size
                print(f"✅ 截图保存成功: {file_size} 字节")
                return screenshot_path
            else:
                print("❌ 截图文件未创建")
                return None

        except Exception as e:
            print(f"❌ 截图失败: {e}")
            traceback.print_exc()
            return None

    def print_screenshot_structure(self, report_dir):
        """打印截图目录结构"""
        if not self.screenshot_root.exists():
            print("📸 截图目录为空")
            return

        print("\n📸 截图目录结构:")
        print("="*50)

        # 打印截图根目录
        rel_path = self.screenshot_root.relative_to(report_dir)
        print(f"{rel_path}/")

        # 打印所有测试文件目录
        test_file_dirs = sorted(self.screenshot_root.iterdir())
        total_screenshots = 0

        for i, test_file_dir in enumerate(test_file_dirs):
            if test_file_dir.is_dir():
                test_file_rel = test_file_dir.relative_to(self.screenshot_root)

                # 判断是否是最后一个元素
                if i == len(test_file_dirs) - 1:
                    prefix = "└── "
                else:
                    prefix = "├── "

                print(f"{prefix}{test_file_rel}/")

                # 打印所有测试函数目录
                test_function_dirs = sorted(test_file_dir.iterdir())
                for j, test_function_dir in enumerate(test_function_dirs):
                    if test_function_dir.is_dir():
                        function_rel = test_function_dir.relative_to(test_file_dir)

                        # 判断是否是最后一个元素
                        if j == len(test_function_dirs) - 1:
                            fprefix = "    └── "
                        else:
                            fprefix = "    ├── "

                        print(f"{fprefix}{function_rel}/")

                        # 打印截图文件
                        screenshot_files = sorted(test_function_dir.glob("*.png"))
                        for k, screenshot_file in enumerate(screenshot_files):
                            if k == len(screenshot_files) - 1:
                                file_prefix = "        └── "
                            else:
                                file_prefix = "        ├── "

                            file_size = screenshot_file.stat().st_size
                            status = "✅" if "PASSED" in screenshot_file.name else "❌"
                            print(f"{file_prefix}{status} {screenshot_file.name} ({file_size} bytes)")
                            total_screenshots += 1

        print(f"\n总计截图: {total_screenshots} 个")

# 全局结果收集器实例
test_collector = TestResultCollector()
screenshot_manager = None

# pytest钩子函数
@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    """pytest钩子，用于在测试报告中添加截图和收集结果"""
    global screenshot_manager

    outcome = yield
    report = outcome.get_result()

    # 只在调用阶段处理
    if report.when == 'call':
        screenshot_path = None
        test_full_name = item.nodeid

        print(f"\n{'='*60}")
        print(f"🔍 处理测试结果: {test_full_name}")
        print(f"🔍 测试状态: {'❌ 失败' if report.failed else '✅ 通过'}")

        try:
            # 获取测试类实例
            test_instance = item.instance

            if test_instance:
                print(f"🔍 测试实例: {type(test_instance).__name__}")

                # 检查是否有helper对象
                has_helper = hasattr(test_instance, 'helper')
                print(f"🔍 是否有helper属性: {has_helper}")

                # 尝试截图
                if screenshot_manager:
                    description = f"测试{'失败' if report.failed else '成功'}: {item.name}"

                    # 使用pyautogui截图
                    screenshot_path = screenshot_manager.take_screenshot(
                        test_full_name, report.failed, description
                    )
                else:
                    print("❌ screenshot_manager 未初始化")

            else:
                print("⚠️  未找到测试实例")

        except Exception as e:
            print(f"❌ 处理截图时出错: {e}")
            traceback.print_exc()

        # 添加测试结果（包含截图路径）
        test_collector.add_result(report, screenshot_path)

        print(f"{'='*60}\n")

@pytest.hookimpl(tryfirst=True)
def pytest_configure(config):
    """pytest配置钩子 - 设置HTML报告路径到时间戳文件夹"""
    global screenshot_manager

    # 确保html插件已加载
    if not config.pluginmanager.hasplugin('html'):
        config.pluginmanager.import_plugin('pytest_html.plugin')

    # 创建时间戳报告目录
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    config.report_dir = Path("reports") / timestamp
    config.report_dir.mkdir(parents=True, exist_ok=True)

    # 初始化截图管理器
    screenshot_manager = ScreenshotManager(config.report_dir)

    # 设置HTML报告路径在时间戳文件夹内
    html_report = config.report_dir / f"test_report_{timestamp}.html"

    # 配置HTML报告选项
    config.option.htmlpath = str(html_report)
    config.option.self_contained_html = True

    print(f"🚀 测试配置完成")
    print(f"📁 报告目录: {config.report_dir}")
    print(f"📄 HTML报告: {html_report.name}")

@pytest.hookimpl(trylast=True)
def pytest_sessionfinish(session, exitstatus):
    """测试会话结束时生成报告"""
    try:
        print(f"\n{'='*60}")
        print(f"🎯 测试会话结束，生成报告...")

        # 生成Excel报告
        excel_file = test_collector.generate_excel_report(session.config.report_dir)

        # 生成JSON报告
        json_file = test_collector.generate_json_report(session.config.report_dir)

        # 打印摘要
        test_collector.print_summary(session.config.report_dir)

        # 打印截图目录结构
        if screenshot_manager:
            screenshot_manager.print_screenshot_structure(session.config.report_dir)

        # 生成报告列表文件
        report_list_file = session.config.report_dir / "reports.txt"
        with open(report_list_file, 'w', encoding='utf-8') as f:
            f.write("=" * 60 + "\n")
            f.write("测试报告清单\n")
            f.write("=" * 60 + "\n\n")

            f.write("📊 生成的测试报告:\n")
            f.write("-" * 40 + "\n")
            if excel_file:
                f.write(f"1. Excel报告: {excel_file.name}\n")
            f.write(f"2. HTML报告: test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html\n")
            if json_file:
                f.write(f"3. JSON报告: {json_file.name}\n")

            f.write("\n📸 截图文件结构:\n")
            f.write("-" * 40 + "\n")

            # 列出所有截图文件
            screenshot_root = session.config.report_dir / "screenshots"
            total_screenshots = 0

            if screenshot_root.exists():
                for test_file_dir in sorted(screenshot_root.iterdir()):
                    if test_file_dir.is_dir():
                        f.write(f"\n📁 {test_file_dir.name}/\n")
                        for test_function_dir in sorted(test_file_dir.iterdir()):
                            if test_function_dir.is_dir():
                                f.write(f"  └── 📂 {test_function_dir.name}/\n")
                                for screenshot in sorted(test_function_dir.glob("*.png")):
                                    file_size = screenshot.stat().st_size
                                    status = "[通过]" if "PASSED" in screenshot.name else "[失败]"
                                    f.write(f"      └── {status} {screenshot.name} ({file_size} bytes)\n")
                                    total_screenshots += 1

            f.write(f"\n📈 统计信息:\n")
            f.write("-" * 40 + "\n")
            f.write(f"总测试用例: {test_collector.test_stats['total']}\n")
            f.write(f"通过: {test_collector.test_stats['passed']}\n")
            f.write(f"失败: {test_collector.test_stats['failed']}\n")
            f.write(f"跳过: {test_collector.test_stats['skipped']}\n")
            f.write(f"错误: {test_collector.test_stats['error']}\n")

            if test_collector.test_stats['total'] > 0:
                pass_rate = test_collector.test_stats['passed'] / test_collector.test_stats['total'] * 100
                f.write(f"通过率: {pass_rate:.2f}%\n")

            f.write(f"总截图数: {total_screenshots}\n")
            f.write("=" * 60 + "\n")

        print(f"📋 报告清单: {report_list_file}")
        print(f"{'='*60}\n")

    except Exception as e:
        print(f"❌ 生成报告时发生错误: {e}")
        traceback.print_exc()