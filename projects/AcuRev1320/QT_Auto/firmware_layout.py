"""固件升级对话框坐标解析 —— 把硬编码绝对坐标改为「相对窗口/对话框锚点」的语义化点击。

坐标体系(参考 comm/ctl_acuview):
  - 坐标外置到 data/firmware_layout.json，存的是「相对锚点(主窗口 / Add Connection 对话框)
    左上角」的相对坐标(物理像素)。
  - 运行期用 pywinauto 取锚点实际 rectangle()，绝对坐标 = 锚点.left/top + 相对坐标。
  - 进程设 DPI 感知(复用 comm/ctl_acuview/dpi)，物理像素一致，避免 125% 缩放下坐标错位。

基线: 1920x1080 / 125% / 主窗口最大化。换分辨率/缩放后用 plan_overlay 核对、必要时重标坐标。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

# 让本文件既能被 pytest 导入、也能直接 `python firmware_layout.py` 运行(标定模式)
_REPO_ROOT = Path(__file__).resolve().parents[3]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from comm.ctl_acuview import dpi  # noqa: F401,E402  导入即设 DPI 感知(须在 pywinauto 之前)


class FirmwareLayoutError(RuntimeError):
    pass


class FirmwareLayout:
    """读 firmware_layout.json，按锚点把相对坐标解析为绝对坐标并点击。"""

    def __init__(self, helper, json_path: str | Path):
        self.helper = helper
        self.path = Path(json_path)
        with open(self.path, "r", encoding="utf-8") as fh:
            self.data = json.load(fh)

    # ---- 锚点矩形(pywinauto 实测) ----
    def _anchor_rect(self, anchor_name: str):
        """返回锚点窗口/对话框的实际矩形(含 .left/.top)。每次实测，避免窗口移动后坐标失效。"""
        a = self.data.get("anchors", {}).get(anchor_name)
        if not a:
            raise FirmwareLayoutError(f"未定义锚点: {anchor_name}")
        from pywinauto import Desktop

        def _area(w):
            r = w.rectangle()
            return (r.right - r.left) * (r.bottom - r.top)

        if a.get("type") == "window":
            # 升级时 Meter Update 弹窗的 OS 标题也叫 "Acuview 2" → 同名窗口可能多于一个,
            # 取面积最大的(=最大化的主窗口),避免 pywinauto 的 ElementAmbiguousError。
            wins = Desktop(backend="uia").windows(title_re=a["title_re"], top_level_only=True)
            if not wins:
                raise FirmwareLayoutError(f"未找到窗口: {a['title_re']}")
            return max(wins, key=_area).rectangle()
        if a.get("type") == "dialog":
            wins = Desktop(backend="uia").windows(title=a["title"], top_level_only=True)
            if not wins:
                raise FirmwareLayoutError(f"未找到对话框: {a['title']}")
            return wins[0].rectangle()
        raise FirmwareLayoutError(f"未知锚点类型: {a.get('type')}")

    # ---- 语义化点击 ----
    def click(self, widget_name: str):
        w = self.data.get("widgets", {}).get(widget_name)
        if not w:
            raise FirmwareLayoutError(f"未定义控件: {widget_name}")
        self._click_rel(w["anchor"], w.get("x"), w.get("y"), widget_name)

    def click_baud(self, rate: str):
        bo = self.data.get("baud_options", {})
        opt = bo.get("options", {}).get(str(rate))
        if not opt:
            raise FirmwareLayoutError(f"未定义波特率选项: {rate}")
        self._click_rel(bo["anchor"], opt.get("x"), opt.get("y"), f"baud:{rate}")

    def click_com(self, com: str):
        """点击 Meter Update 列表中指定 COM 口(如 "11")的 Select 复选框。

        坐标按 COM 号标定(见 firmware_layout.json 的 com_checkboxes)；扫描列表在
        本台架稳定，故按号存固定相对坐标，换台架/换串口重标对应项即可。
        """
        cc = self.data.get("com_checkboxes", {})
        opt = cc.get("options", {}).get(str(com))
        if not opt:
            raise FirmwareLayoutError(f"未定义 COM 口复选框: {com}(请在 {self.path.name} 的 com_checkboxes 标定)")
        self._click_rel(cc["anchor"], opt.get("x"), opt.get("y"), f"com:{com}")

    def _click_rel(self, anchor: str, rx, ry, label: str):
        if rx is None or ry is None:
            raise FirmwareLayoutError(
                f"控件 {label} 坐标未标定(x/y 为空)。请在基线真机上标定后写入 {self.path.name}。")
        rect = self._anchor_rect(anchor)
        self.helper.click_rel(rect, rx, ry)
        self.helper.logger.info(f"点击控件 {label}: 锚点 {anchor} + 相对({rx},{ry})")

    # ---- 标定辅助: 把当前各控件解析点画在截图上, 供肉眼核对(安全, 不点击) ----
    def plan_overlay(self, out_path: str | Path) -> str:
        """在当前屏幕截图上标注每个控件的解析点击点，用于真跑前核对坐标映射。"""
        import pyautogui
        from PIL import ImageDraw
        img = pyautogui.screenshot()
        draw = ImageDraw.Draw(img)
        for name, w in self.data.get("widgets", {}).items():
            if w.get("x") is None or w.get("y") is None:
                continue
            try:
                rect = self._anchor_rect(w["anchor"])
            except FirmwareLayoutError:
                continue
            cx, cy = rect.left + int(w["x"]), rect.top + int(w["y"])
            draw.ellipse([cx - 5, cy - 5, cx + 5, cy + 5], outline=(255, 60, 60), width=2)
            draw.text((cx + 7, cy - 7), name, fill=(255, 60, 60))
        out = str(out_path)
        img.save(out)
        return out


# ---------------------------------------------------------------------------
# 标定 / 核对入口: 直接 `python projects/AcuRev1320/QT_Auto/firmware_layout.py`
#   无参数  : 实时打印鼠标绝对坐标 + 相对各锚点的偏移(把鼠标移到控件上读数填 JSON)
#   --overlay OUT.png : 把各控件解析点画在当前截图上保存,供肉眼核对坐标映射
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import argparse
    import time

    import pyautogui

    _here = Path(__file__).resolve().parent
    _json = _here / "data" / "firmware_layout.json"

    _parser = argparse.ArgumentParser(description="固件升级坐标标定/核对工具")
    _parser.add_argument("--overlay", metavar="OUT",
                         help="把各控件解析点画在截图上保存到该路径(如 check.png)")
    _args = _parser.parse_args()

    _layout = FirmwareLayout(helper=None, json_path=_json)

    if _args.overlay:
        print(f"已保存核对图: {_layout.plan_overlay(_args.overlay)}")
        sys.exit(0)

    # 实时坐标模式: 先连各锚点拿矩形, 再循环打印鼠标相对偏移
    _anchors = {}
    for _name in _layout.data.get("anchors", {}):
        try:
            _r = _layout._anchor_rect(_name)
            _anchors[_name] = _r
            print(f"[锚点] {_name:<10} rect=({_r.left},{_r.top})-({_r.right},{_r.bottom})  "
                  f"尺寸 {_r.right - _r.left}x{_r.bottom - _r.top}")
        except Exception as _e:  # noqa: BLE001  标定工具,任何连接失败都只提示
            print(f"[锚点] {_name:<10} 不可用(对应窗口/对话框未打开?): {_e}")

    if not _anchors:
        print("没有可用锚点。请先打开 Acuview 2 / 对应对话框后重试。")
        sys.exit(1)

    print("\n把鼠标移到目标控件上, 读「相对偏移」填进 firmware_layout.json。Ctrl+C 退出。\n")
    try:
        while True:
            _x, _y = pyautogui.position()
            _rels = "  ".join(f"{_n}:({_x - _r.left},{_y - _r.top})" for _n, _r in _anchors.items())
            print(f"鼠标绝对({_x},{_y})  相对偏移 {_rels}          ", end="\r")
            time.sleep(0.2)
    except KeyboardInterrupt:
        print("\n已退出。")
