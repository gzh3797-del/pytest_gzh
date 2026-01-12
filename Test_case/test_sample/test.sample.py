from pywinauto import Desktop
import subprocess
import time

# 1. 使用系统命令启动记事本（不依赖 Pywinauto 的进程绑定）
subprocess.Popen("notepad.exe")
time.sleep(2)  # 等两秒，确保窗口弹出来了

# 2. 使用 Desktop 连接，它会在所有打开的软件里找
# 注意：我在正则里同时写了中文和英文，确保能匹配上
desktop = Desktop(backend="uia")
dlg = desktop.window(title_re=".*记事本.*|.*Notepad.*")

# 3. 显式等待窗口出现（防止电脑卡顿导致报错）
dlg.wait('visible', timeout=10)

# 4. 打印控件结构（这一步能成功，说明窗口找到了）
print("窗口已找到！正在打印控件结构...")
# dlg.print_control_identifiers()  # 调试时可以取消注释这行

# 5. 操作输入框
# Windows 11 记事本的编辑区通常叫 "RichEditD2DPT" 或简单地找 "Edit" 类型
# 我们用 best_match 让它自己去模糊匹配最像编辑区的地方
try:
    # 尝试直接向窗口发送按键（最通用的方法）
    dlg.type_keys("Hello Gemini!", with_spaces=True)
except Exception as e:
    print(f"输入失败: {e}")