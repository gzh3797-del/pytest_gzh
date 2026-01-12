from pywinauto import Desktop

try:
    # 1. 连接到 Acuview 2 窗口
    # backend="uia" 对现代界面支持更好，如果报错或找不到，下文有备选方案
    app_window = Desktop(backend="uia").window(title_re=".*Acuview 2.*")

    # 等待窗口可见，防止报错
    app_window.wait('visible', timeout=10)

    # 2. 打印控件树结构
    print("正在分析界面，请稍候...")
    app_window.print_control_identifiers()

except Exception as e:
    print(f"发生错误: {e}")