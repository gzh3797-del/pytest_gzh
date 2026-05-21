import pyautogui
import time

def launch_acuview(icon_path, wait_seconds=3):
    """
    模拟双击桌面图标启动 Acuview 2

    icon_path: 桌面图标截图路径
    wait_seconds: 启动前等待时间
    """
    # print(f"请确保桌面显示并图标截图路径正确: {icon_path}")
    time.sleep(wait_seconds)  # 给用户切换到桌面时间

    # 尝试查找图标
    icon_location = pyautogui.locateOnScreen(icon_path, confidence=0.9)
    if icon_location:
        center = pyautogui.center(icon_location)
        print(f"找到图标位置：{center}, 正在双击...")
        pyautogui.doubleClick(center)
        print("Acuview 2 已启动（已双击图标）")
    else:
        print("未找到桌面图标，请确认截图正确或图标可见")

# -------------------------
# 使用示例
# -------------------------
if __name__ == "__main__":
    # 先截取桌面 Acuview 2 图标，保存为 acuview_icon.png
    icon_screenshot_path = "png/acuview_icon.png"  # 修改为你保存的图标截图路径
    launch_acuview(icon_screenshot_path)
