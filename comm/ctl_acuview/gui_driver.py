"""驱动 Acuview 2 GUI —— 模拟测试人员操作。

背景(由 uia_probe 实测确认):
  Acuview 2 是 Qt5 自绘界面，UIA 只暴露窗口外框，内部控件(树/输入框/下拉/开关)全部不可见。
  => 必须用 *坐标 + 视觉(截图/模板匹配/OCR)* 方式驱动。

坐标体系:
  - 进程设为 DPI 感知(物理像素 1920x1080)，截图/点击/窗口矩形三者一致。
  - JSON 控件坐标是相对"页面内容区"的相对坐标(x,y)。
  - 内容区原点(content_origin) = 窗口客户区原点 + 标定插入量(insets)，由 calibrate 标定后写入 config。
  - 控件绝对坐标 = content_origin + (widget.x, widget.y) [- 垂直滚动量]。

安全:
  - 锁屏/安全桌面下拒绝点击(截图与点击都不可靠)。
  - 写操作的护栏在 verify/demo 层(read→write→verify→restore)。
"""
from __future__ import annotations

import ctypes
import re
import time
from dataclasses import dataclass

from . import dpi  # noqa: F401  导入即设 DPI 感知(必须在 pyautogui 之前)
import pyautogui
from PIL import Image, ImageOps

from .config import get_config
from .spec_loader import load_spec

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1


# --------------------------------------------------------------------------
# 会话锁屏检测
# --------------------------------------------------------------------------
def _logonui_running() -> bool:
    """枚举进程判断 LogonUI.exe 是否存在(Win11 锁屏/安全桌面时存在)。快速、无子进程。"""
    TH32CS_SNAPPROCESS = 0x00000002
    INVALID = ctypes.c_void_p(-1).value
    k32 = ctypes.windll.kernel32

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", ctypes.c_ulong), ("cntUsage", ctypes.c_ulong),
            ("th32ProcessID", ctypes.c_ulong), ("th32DefaultHeapID", ctypes.c_void_p),
            ("th32ModuleID", ctypes.c_ulong), ("cntThreads", ctypes.c_ulong),
            ("th32ParentProcessID", ctypes.c_ulong), ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", ctypes.c_ulong), ("szExeFile", ctypes.c_wchar * 260),
        ]

    snap = k32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snap == INVALID:
        return False
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        if not k32.Process32FirstW(snap, ctypes.byref(entry)):
            return False
        while True:
            if entry.szExeFile.lower() == "logonui.exe":
                return True
            if not k32.Process32NextW(snap, ctypes.byref(entry)):
                break
        return False
    finally:
        k32.CloseHandle(snap)


def is_session_locked() -> bool:
    """会话是否锁屏/处于安全桌面。

    用两条信号取或：
      1) LogonUI.exe 存在(Win11 锁屏时常驻) —— 实测对本机可靠；
      2) 无法打开输入桌面 —— 安全桌面下用户进程无权访问。
    截图/点击在锁屏下不可靠，故宁可误报为锁定。
    """
    if _logonui_running():
        return True
    user32 = ctypes.windll.user32
    hdesk = user32.OpenInputDesktop(0, False, 0x0001)  # DESKTOP_READOBJECTS
    if not hdesk:
        return True
    user32.CloseDesktop(hdesk)
    return False


# --------------------------------------------------------------------------
@dataclass
class Rect:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self):
        return self.right - self.left

    @property
    def height(self):
        return self.bottom - self.top


class GuiError(RuntimeError):
    pass


class AppDriver:
    """Acuview 2 窗口驱动 + 坐标/视觉操作原语。"""

    # 会话内主窗口缓存: _find_main_window 走 UIA 全桌面枚举, portable debug 多同名窗口时单次
    # 可达 ~20s。同一进程内首条用例定位后缓存句柄, 后续用例直接复用(20s → ~0s), 是批跑提速关键。
    _cached_win = None

    def __init__(self):
        self.cfg = get_config()
        self.registers, self.pages = load_spec()
        self.app = None
        self.win = None
        self.reused = False   # 本次 launch_or_connect 是否复用了缓存窗口(供 _connect 跳过多余步骤)
        self._tess_ready = False
        self._configure_tesseract()

    # ---- Tesseract ----
    def _configure_tesseract(self):
        cmd = (self.cfg.gui_backend.get("tesseract_cmd") or "").strip()
        try:
            import pytesseract
            if cmd:
                pytesseract.pytesseract.tesseract_cmd = cmd
            self._pytesseract = pytesseract
        except Exception:
            self._pytesseract = None

    # 主窗口标题过滤: 排除同样匹配 window_title_re 的干扰窗口
    #   (Add Connection 对话框 / 打开了 portable 目录的文件资源管理器 / Program Manager 等)。
    _MAIN_WIN_EXCLUDE = ("Add Connection", "文件资源管理器", "File Explorer",
                         "Debug_100 -", "Program Manager")

    def _find_main_window(self, timeout: float = 20.0):
        """在所有匹配 window_title_re 的可见窗口里挑出 Acuview 2 主窗口。

        portable debug 构建一次启动会拉起多个同名窗口, 且用户可能打开了名字含
        "Acuview 2" 的资源管理器 → Application.connect(title_re=...) 会抛
        ElementAmbiguousError。这里按"排除干扰标题 + 取可见面积最大者"稳健定位。
        """
        from pywinauto import Desktop
        title_re = self.cfg.app.window_title_re
        deadline = time.time() + timeout
        while time.time() < deadline:
            cands = []
            for w in Desktop(backend="uia").windows(title_re=title_re, visible_only=True):
                try:
                    txt = w.window_text()
                    if any(bad in txt for bad in self._MAIN_WIN_EXCLUDE):
                        continue
                    r = w.rectangle()
                    cands.append((r.width() * r.height(), w))
                except Exception:  # noqa: BLE001
                    continue
            if cands:
                cands.sort(key=lambda x: -x[0])
                return cands[0][1]
            time.sleep(0.5)
        return None

    # ---- 启动 / 连接 ----
    def launch_or_connect(self, require_unlocked: bool = True):
        if require_unlocked and is_session_locked():
            raise GuiError("当前为锁屏/安全桌面，无法进行界面截图与点击。请先解锁该机器再运行 GUI 演示。")
        from pywinauto import Application
        # 复用会话内已定位的主窗口(仍有效则跳过昂贵的 UIA 全桌面枚举)。
        cached = AppDriver._cached_win
        if cached is not None:
            try:
                if cached.rectangle().width() > 0:
                    self.win = cached
                    self.reused = True
                    try:
                        self.win.set_focus()
                    except Exception:
                        pass
                    return self
            except Exception:
                AppDriver._cached_win = None
        win = self._find_main_window(timeout=3)
        if win is None:
            # 关键: 必须在安装目录下启动，否则 Add Connection 等对话框按相对路径加载内容会渲染空白。
            import os
            from .app_locator import resolve_acuview_exe
            exe = resolve_acuview_exe((self.cfg.app.get("exe_path") or "").strip())
            if not exe:
                raise GuiError(
                    "未找到 Acuview 2.exe，且无运行实例可连接。请在 config.yaml 的 "
                    "app.exe_path 填写绝对路径(或确认 Acuview 2 已按默认方式安装)。")
            workdir = os.path.dirname(exe)
            self.app = Application(backend="uia").start(f'"{exe}"', work_dir=workdir, timeout=10)
            time.sleep(self.cfg.app.launch_wait_s)
            win = self._find_main_window(timeout=20)
        if win is None:
            raise GuiError("未能定位 Acuview 2 主窗口(多个同名窗口干扰且过滤后为空)")
        self.win = win
        AppDriver._cached_win = win   # 缓存供同会话后续用例复用
        self.reused = False
        try:
            self.win.set_focus()
        except Exception:
            pass
        return self

    def window_rect(self) -> Rect:
        r = self.win.rectangle()
        return Rect(r.left, r.top, r.right, r.bottom)

    # ---- 坐标映射 ----
    def content_origin(self) -> tuple[int, int]:
        """页面内容区左上角的屏幕坐标。

        优先用 config 标定的 fixed 偏移(相对窗口客户区)；template 方式见 calibrate.py。
        """
        rect = self.window_rect()
        co = self.cfg.gui_backend.content_origin
        if co.get("method") == "fixed":
            return rect.left + int(co.get("fixed_x", 0)), rect.top + int(co.get("fixed_y", 0))
        # 默认猜测(需 calibrate 校准): 左侧导航树约 250px，顶部标题/工具栏约 150px
        return rect.left + 250, rect.top + 150

    def widget_abs(self, page: str, widget_name: str, scroll_y: int = 0) -> tuple[int, int]:
        w = self._widget(page, widget_name)
        ox, oy = self.content_origin()
        cx = ox + int(w["x"]) + int(w.get("w", 0)) // 2
        cy = oy + int(w["y"]) - scroll_y + int(w.get("h", 0)) // 2
        return cx, cy

    def _widget(self, page: str, widget_name: str) -> dict:
        p = self.pages["pages"].get(page)
        if not p:
            raise GuiError(f"页面不存在: {page}")
        w = p["widgets"].get(widget_name)
        if not w:
            raise GuiError(f"页面 {page} 无控件: {widget_name}")
        if w.get("x") is None or w.get("y") is None:
            raise GuiError(f"控件 {widget_name} 无坐标")
        return w

    # ---- 截图 / OCR ----
    def screenshot(self, region: tuple | None = None):
        return pyautogui.screenshot(region=region)

    def ocr_words(self, region: tuple):
        """region=(x,y,w,h) 物理像素; 返回 [{text,left,top,width,height,conf}](绝对坐标)。"""
        if not self._pytesseract:
            raise GuiError("pytesseract 不可用，无法 OCR")
        x, y, w, h = region
        img = pyautogui.screenshot(region=(x, y, w, h))
        data = self._pytesseract.image_to_data(img, output_type=self._pytesseract.Output.DICT)
        out = []
        for i, txt in enumerate(data["text"]):
            if txt.strip():
                out.append({
                    "text": txt.strip(),
                    "left": x + data["left"][i], "top": y + data["top"][i],
                    "width": data["width"][i], "height": data["height"][i],
                    "conf": float(data["conf"][i]),
                })
        return out

    def ocr_text(self, region: tuple) -> str:
        if not self._pytesseract:
            raise GuiError("pytesseract 不可用，无法 OCR")
        x, y, w, h = region
        img = pyautogui.screenshot(region=(x, y, w, h))
        return self._pytesseract.image_to_string(img).strip()

    # ---- 点击原语(带锁屏护栏) ----
    def _guard(self):
        if is_session_locked():
            raise GuiError("会话已锁屏，已阻止点击操作。")

    def click_abs(self, x: int, y: int, double: bool = False):
        self._guard()
        pyautogui.moveTo(x, y, duration=0.1)
        if double:
            pyautogui.doubleClick()
        else:
            pyautogui.click()

    def type_text(self, text: str, clear: bool = True):
        self._guard()
        if clear:
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("delete")
        pyautogui.typewrite(str(text), interval=0.03)

    # ---- 导航(左侧树, OCR 定位文本) ----
    def navigate(self, item_text: str, panel_region: tuple | None = None) -> bool:
        """在左侧导航树用 OCR 找到 item_text 并点击。"""
        rect = self.window_rect()
        if panel_region is None:
            panel_region = (rect.left, rect.top + 120, 260, rect.height - 140)
        words = self.ocr_words(panel_region)
        target = item_text.strip().lower()
        # 整词或包含匹配
        for wd in sorted(words, key=lambda d: -d["conf"]):
            if target == wd["text"].lower() or target in wd["text"].lower():
                cx = wd["left"] + wd["width"] // 2
                cy = wd["top"] + wd["height"] // 2
                self.click_abs(cx, cy)
                time.sleep(0.6)
                return True
        return False

    def navigate_path(self, path: list[str]) -> bool:
        ok = True
        for item in path:
            ok = self.navigate(item) and ok
            time.sleep(0.4)
        return ok

    # ---- 设值(按控件类型) ----
    def set_value(self, page: str, widget_name: str, value, scroll_y: int = 0):
        w = self._widget(page, widget_name)
        wt = w["type"]
        x, y = self.widget_abs(page, widget_name, scroll_y)
        if wt in ("spinBox", "lineEdit", "ipEdit", "doubleSpinBox"):
            self.click_abs(x, y)
            self.type_text(value, clear=True)
            pyautogui.press("tab")
        elif wt == "comboBox":
            self.click_abs(x, y)
            time.sleep(0.5)
            # 下拉弹层与控件*左边缘*对齐、在控件*下方*展开; 以控件几何(非中心点)为锚。
            ox, oy = self.content_origin()
            left = ox + int(w["x"])
            top = oy + int(w["y"]) - scroll_y
            ww, hh = int(w.get("w", 130)), int(w.get("h", 30))
            popup = (left - 5, top + hh, ww + 60, 320)
            self._select_combo_option(str(value), popup)
        elif wt == "switchButton":
            self._set_switch(page, widget_name, bool_value=bool(value), pos=(x, y))
        else:
            raise GuiError(f"暂不支持设值的控件类型: {wt}")
        time.sleep(0.3)

    def _select_combo_option(self, value: str, popup: tuple):
        """下拉展开后, 在 popup 区域(=控件左边缘/下方)增强 OCR 找选项文本并点击。

        此 Acuview 版本下拉字号小、默认 OCR 读不出 → 灰度 + 3x 放大 + psm6;
        数字型(波特率)加数字白名单(纠正 9→$ 类误识)。当前值行(蓝底白字高亮)读不到,
        但选值恒为"非当前值", 不影响命中。
        """
        numeric = str(value).strip().replace(".", "").isdigit()
        words = self._combo_popup_words(popup, numeric=numeric)
        if numeric:
            tgt = re.sub(r"\D", "", str(value))
            match = [wd for wd in words if re.sub(r"\D", "", wd["text"]) == tgt]
        else:
            tgt = str(value).strip().lower()
            match = [wd for wd in words
                     if tgt in wd["text"].lower() or wd["text"].lower() in tgt]
        if match:
            self.click_abs(int(match[0]["cx"]), int(match[0]["cy"]))
            time.sleep(0.4)
            return True
        # 无匹配但下拉已读到其它选项: OCR 恒读不到"当前高亮行", 故目标必为当前值 → 已选中, 免点。
        # (回读走 Modbus 权威判决, 该启发式最坏只会误判 FAIL, 不会误 PASS。)
        if words:
            pyautogui.press("esc")
            time.sleep(0.3)
            return True
        raise GuiError(f"下拉未读到任何选项(未展开/OCR失效): {value}")

    def _combo_popup_words(self, region: tuple, numeric: bool = False) -> list[dict]:
        """下拉弹层增强 OCR: 灰度 + 3x 放大 + psm6(数字型加白名单)。
        返回 [{text, cx, cy, conf}], cx/cy 为绝对屏幕坐标(已把放大坐标折回)。
        """
        if not self._pytesseract:
            raise GuiError("pytesseract 不可用，无法 OCR")
        x, y, w, h = region
        up = 3
        img = ImageOps.grayscale(pyautogui.screenshot(region=(x, y, w, h)))
        resample = getattr(Image, "Resampling", Image).LANCZOS
        img = img.resize((img.width * up, img.height * up), resample)
        cfg = "--psm 6"
        if numeric:
            cfg += " -c tessedit_char_whitelist=0123456789"
        data = self._pytesseract.image_to_data(
            img, config=cfg, output_type=self._pytesseract.Output.DICT)
        out = []
        for i, txt in enumerate(data["text"]):
            if txt.strip():
                out.append({
                    "text": txt.strip(),
                    "cx": x + (data["left"][i] + data["width"][i] / 2) / up,
                    "cy": y + (data["top"][i] + data["height"][i] / 2) / up,
                    "conf": float(data["conf"][i]),
                })
        return out

    def _set_switch(self, page: str, widget_name: str, bool_value: bool, pos: tuple):
        cur = self.get_switch_state(page, widget_name)
        if cur is None or cur != bool_value:
            self.click_abs(*pos)

    # ---- 取值(OCR) ----
    def get_value(self, page: str, widget_name: str, scroll_y: int = 0) -> str:
        w = self._widget(page, widget_name)
        ox, oy = self.content_origin()
        x = ox + int(w["x"])
        y = oy + int(w["y"]) - scroll_y
        region = (x, y, max(60, int(w.get("w", 80))), max(20, int(w.get("h", 30))))
        return self.ocr_text(region)

    def get_switch_state(self, page: str, widget_name: str, scroll_y: int = 0):
        txt = self.get_value(page, widget_name, scroll_y).upper()
        if "ON" in txt:
            return True
        if "OFF" in txt:
            return False
        return None

    # ---- 保存 / 读取按钮(OCR 定位) ----
    def click_button_by_text(self, text: str, search_region: tuple | None = None) -> bool:
        rect = self.window_rect()
        if search_region is None:
            search_region = (rect.left, rect.top, rect.width, rect.height)
        for wd in self.ocr_words(search_region):
            if text.lower() in wd["text"].lower():
                self.click_abs(wd["left"] + wd["width"] // 2, wd["top"] + wd["height"] // 2)
                time.sleep(0.8)
                return True
        return False

    def click_save(self):
        return (self.click_button_by_text("Save") or self.click_button_by_text("Apply")
                or self.click_button_by_text("保存"))

    def click_read(self):
        return (self.click_button_by_text("Read") or self.click_button_by_text("Refresh")
                or self.click_button_by_text("读取"))

    # ---- 模板匹配(OpenCV, 无需 OCR 引擎)定位并点击按钮/对话框 ----
    def find_template(self, template_path: str, region: tuple | None = None,
                      threshold: float = 0.82):
        import cv2
        import numpy as np
        shot = pyautogui.screenshot(region=region)
        hay = cv2.cvtColor(np.array(shot), cv2.COLOR_RGB2BGR)
        # cv2.imread 不支持中文路径(本项目位于 C:\AI工具\...), 改用 np.fromfile + imdecode
        try:
            tmpl = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        except (FileNotFoundError, OSError):
            tmpl = None
        if tmpl is None:
            raise GuiError(f"模板不存在或无法解码: {template_path}")
        res = cv2.matchTemplate(hay, tmpl, cv2.TM_CCOEFF_NORMED)
        _, maxv, _, maxloc = cv2.minMaxLoc(res)
        if maxv < threshold:
            return None
        th, tw = tmpl.shape[:2]
        ox, oy = (region[0], region[1]) if region else (0, 0)
        return (ox + maxloc[0] + tw // 2, oy + maxloc[1] + th // 2, float(maxv))

    def click_template(self, template_path: str, threshold: float = 0.82,
                       region: tuple | None = None) -> bool:
        hit = self.find_template(template_path, region, threshold)
        if hit:
            self.click_abs(hit[0], hit[1])
            return True
        return False

    # ---- 窗口最大化(坐标稳定): 校验宽度并重试, 否则后续模板/坐标全部错位 ----
    def maximize(self):
        for _ in range(4):
            try:
                pyautogui.press("esc")          # 关掉可能打开的菜单
                self.win.set_focus()
                self.win.maximize()
            except Exception:
                pass
            time.sleep(1.5)
            try:
                r = self.window_rect()
                if r.width >= 1900:
                    return r
            except Exception:
                pass
        return self.window_rect()

    # ---- 连接管理: 在 Add Connection 对话框选一行并连接 ----
    # Add Connection 对话框内的相对坐标(相对对话框左上角)。对话框位置每次启动会变,
    # 故必须以对话框实际矩形为锚, 不能用屏幕绝对坐标。
    CONN_CHECKBOX_DX = 47      # Select 复选框列中心
    CONN_ROW0_DY = 177         # 第 0 行中心
    CONN_ROW_STEP = 40         # 行高
    CONN_CONNECT_DX = 858      # Connect 按钮中心
    CONN_CONNECT_DY = 74

    def connect_meter(self, row_index: int = 0) -> bool:
        """选择已保存的连接并点 Connect。坐标锚定到 Add Connection 对话框实际矩形。

        无 OCR 时按行号选(默认第 0 行=本测试的 192.168.3.37/AcuRev1320);
        装了 Tesseract 后可改为按 IP/型号匹配行(见 README)。
        """
        from pywinauto import Application
        app = Application(backend="uia").connect(title_re="Add Connection", timeout=10)
        w = app.window(title="Add Connection")
        w.set_focus()
        time.sleep(0.6)
        r = w.rectangle()
        # 先点标题栏激活窗口: 未激活窗口的首个点击会被"激活"吞掉, 复选框点不上
        self.click_abs(r.left + 400, r.top + 12)
        time.sleep(0.3)
        self.click_abs(r.left + self.CONN_CHECKBOX_DX,
                       r.top + self.CONN_ROW0_DY + row_index * self.CONN_ROW_STEP)
        time.sleep(0.6)
        self.click_abs(r.left + self.CONN_CONNECT_DX, r.top + self.CONN_CONNECT_DY)
        time.sleep(5)
        return True

    # ---- 下发并确认(处理 密码 + "Do you want to update?" 对话框) ----
    def update_and_confirm(self, update_xy: tuple, password: str = "0000",
                           timeout: float = 9.0) -> bool:
        """点 Update，按出现顺序处理对话框：
           1) 若需权限，弹"Please Enter Password"(预填 0000) -> Confirm
           2) "Do you want to update?" -> Yes
        用模板匹配 Confirm/Yes 按钮，避免盲点击竞态。
        """
        import os
        tdir = os.path.join(os.path.dirname(__file__), "..", "templates_acuview")
        confirm_t = os.path.join(tdir, "btn_confirm.png")
        yes_t = os.path.join(tdir, "btn_yes.png")
        self.click_abs(*update_xy)
        deadline = time.time() + timeout
        did_confirm = False
        yes_clicks = 0      # 依次点掉: "Do you want to update?" 的 Yes + "Update Successful!" 结果框按钮
        while time.time() < deadline:
            time.sleep(0.5)
            # 弹窗密码框(权限不足时)用高阈值, 避免误匹配 "Do you want to update?" 的 Yes
            if not did_confirm and self.click_template(confirm_t, threshold=0.95):
                did_confirm = True
                time.sleep(0.8)
                continue
            if self.click_template(yes_t, threshold=0.9):
                yes_clicks += 1
                time.sleep(1.0)
                continue
            if yes_clicks >= 1:   # 已无可点的蓝色按钮 -> 确认框与结果框都已关闭
                break
        time.sleep(0.7)
        return yes_clicks >= 1

    # ---- 可视化:把计划点击点画在截图上(安全,不点击) ----
    def plan_overlay(self, page: str, out_path: str | None = None, scroll_y: int = 0) -> str:
        """在当前截图上标注该页所有控件的计划点击点，用于人工核对坐标映射。"""
        from PIL import ImageDraw
        img = pyautogui.screenshot()
        draw = ImageDraw.Draw(img)
        p = self.pages["pages"].get(page, {})
        ox, oy = self.content_origin()
        for name, w in p.get("widgets", {}).items():
            if w.get("x") is None:
                continue
            cx = ox + int(w["x"]) + int(w.get("w", 0)) // 2
            cy = oy + int(w["y"]) - scroll_y + int(w.get("h", 0)) // 2
            color = (255, 60, 60) if "Value" in name or "Edit" in name or "Spin" in name or "Combo" in name else (80, 160, 255)
            draw.ellipse([cx - 4, cy - 4, cx + 4, cy + 4], outline=color, width=2)
            draw.text((cx + 6, cy - 6), name[:24], fill=color)
        self.cfg.report_dir.mkdir(parents=True, exist_ok=True)
        out = out_path or str(self.cfg.report_dir / f"plan_{page}.png")
        img.save(out)
        return out

    def close(self):
        pass  # 默认保留窗口供观察；如需关闭可在此实现


if __name__ == "__main__":
    print("锁屏状态:", is_session_locked())
    d = AppDriver()
    print(f"已加载页面 {len(d.pages['pages'])} 个; 寄存器 {d.registers['count']} 条")
    print("Communication 页输入类控件:")
    for n, w in d.pages["pages"]["Communication"]["widgets"].items():
        if w["type"] in ("spinBox", "comboBox", "ipEdit", "switchButton", "lineEdit"):
            print(f"  {n:<46} type={w['type']:<12} rel=({w['x']},{w['y']})")
