from pywinauto import Desktop

# 列出桌面上所有窗口的标题
windows = Desktop(backend="uia").windows()
for w in windows:
    print(w.window_text())

