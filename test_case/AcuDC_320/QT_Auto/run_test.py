#!/usr/bin/env python3
import pytest
import sys
import os
import ctypes
import subprocess
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


class WindowsKeepAwake:
    """Windows系统防睡眠工具"""

    def __init__(self):
        self.is_windows = sys.platform == "win32"
        self.active = False
        if self.is_windows:
            self.ES_CONTINUOUS = 0x80000000
            self.ES_SYSTEM_REQUIRED = 0x00000001
            self.ES_DISPLAY_REQUIRED = 0x00000002

    def __enter__(self):
        if self.is_windows:
            self.prevent_sleep()
            self.disable_screen_timeout()
            self.active = True
            print("🔋 防睡眠已启用 - 系统不会睡眠或锁屏")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.is_windows and self.active:
            self.restore_sleep()
            self.restore_screen_timeout()
            self.active = False
            print("🔋 防睡眠已禁用 - 恢复系统正常电源管理")

    def prevent_sleep(self):
        """防止系统睡眠"""
        ctypes.windll.kernel32.SetThreadExecutionState(
            self.ES_CONTINUOUS | self.ES_SYSTEM_REQUIRED | self.ES_DISPLAY_REQUIRED
        )

    def restore_sleep(self):
        """恢复系统睡眠"""
        ctypes.windll.kernel32.SetThreadExecutionState(self.ES_CONTINUOUS)

    def disable_screen_timeout(self):
        """禁用屏幕超时"""
        try:
            if sys.platform == "win32":
                # Windows: 使用powercfg命令
                subprocess.run(["powercfg", "/change", "monitor-timeout-ac", "0"],
                               capture_output=True, shell=True)
                subprocess.run(["powercfg", "/change", "standby-timeout-ac", "0"],
                               capture_output=True, shell=True)
                print("💡 已禁用显示器和系统睡眠超时")
            elif sys.platform == "darwin":  # macOS
                subprocess.Popen(["caffeinate", "-dim"])
                print("🍎 macOS: 已启用caffeinate防止睡眠")
            elif sys.platform == "linux":
                # Linux: 使用xset命令
                subprocess.run(["xset", "s", "off"], capture_output=True)
                subprocess.run(["xset", "-dpms"], capture_output=True)
                print("🐧 Linux: 已禁用屏幕保护和DPMS")
        except Exception as e:
            print(f"⚠️ 无法修改电源设置: {e}")

    def restore_screen_timeout(self):
        """恢复屏幕超时设置"""
        try:
            if sys.platform == "win32":
                # 恢复Windows默认设置
                subprocess.run(["powercfg", "/change", "monitor-timeout-ac", "15"],
                               capture_output=True, shell=True)
                subprocess.run(["powercfg", "/change", "standby-timeout-ac", "30"],
                               capture_output=True, shell=True)
                print("💡 已恢复显示器和系统睡眠超时设置")
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["pkill", "caffeinate"], capture_output=True)
                print("🍎 macOS: 已停止caffeinate")
            elif sys.platform == "linux":
                subprocess.run(["xset", "s", "on"], capture_output=True)
                subprocess.run(["xset", "+dpms"], capture_output=True)
                print("🐧 Linux: 已恢复屏幕保护和DPMS")
        except Exception as e:
            print(f"⚠️ 无法恢复电源设置: {e}")


def run_test_group(pytest_args, group_name="测试组"):
    """运行一组测试"""
    print(f"\n{'=' * 60}")
    print(f"🚀 开始运行 {group_name}")
    print(f"📋 参数: {pytest_args}")
    print(f"{'=' * 60}")

    result = pytest.main(pytest_args)
    success = (result == 0)

    print(f"\n{'=' * 60}")
    print(f"✅ {group_name} {'成功' if success else '失败'}")
    print(f"{'=' * 60}\n")

    return success


def main():
    """主函数"""
    print(f"\n{'=' * 60}")
    print("🎯 开始自动化测试")
    print(f"📁 工作目录: {os.getcwd()}")
    print(f"🐍 Python版本: {sys.version}")
    print(f"🖥️  操作系统: {sys.platform}")
    print(f"{'=' * 60}\n")

    # 启动防睡眠
    with WindowsKeepAwake():
        # 第一组测试
        group1_success = run_test_group([
            # "test_transaction/test_run_stress_charge.py",
            # "test_transaction/test_transaction_config.py",
            "test_time/test_time.py",
            "test_transaction/test_transaction.py",
            "test_echilog/test_echilog.py",
            "test_RTU_communication/test_RTU_communication.py",
            "-v", "--tb=short", "-s", "--color=yes"
        ], "基础功能测试组")

        # 第二组测试
        group2_success = run_test_group([
            "test_firmware/test_firmware.py",
            "-v", "--tb=short", "-s", "-x", "--color=yes"
        ], "固件测试组")

    # 返回最终结果：任意一组失败就返回非0
    final_success = group1_success and group2_success

    print(f"\n{'=' * 60}")
    print(f"🏁 测试完成")
    print(f"📊 最终结果: {'✅ 全部通过' if final_success else '❌ 有测试失败'}")
    print(f"{'=' * 60}")

    return 0 if final_success else 1


if __name__ == "__main__":
    sys.exit(main())