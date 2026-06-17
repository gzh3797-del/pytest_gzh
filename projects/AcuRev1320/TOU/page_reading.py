import time
import pyautogui
import pytesseract
from pywinauto import Application


def assert_page_contains(
    texts,
    *,
    exe_path=None,
    window_title_re=None,
    ocr_region=None,
    tesseract_path=None,
    timeout=10,
):
    """
    断言当前页面包含指定文本（pywinauto 优先，OCR 兜底）

    :param texts: list[str] | str  需要检查的文本
    :param exe_path: 应用 exe 路径（用于 pywinauto）
    :param window_title_re: 窗口标题正则（用于 pywinauto）
    :param ocr_region: (x, y, w, h) OCR 截图区域，None=全屏
    :param tesseract_path: tesseract.exe 路径
    :param timeout: 最大等待时间（秒）
    """

    if isinstance(texts, str):
        texts = [texts]

    end_time = time.time() + timeout
    last_error = ""

    # ---------- 1️⃣ pywinauto ----------
    if exe_path and window_title_re:
        try:
            app = Application(backend="uia").connect(path=exe_path)
            win = app.window(title_re=window_title_re)

            while time.time() < end_time:
                ui_text = " ".join(win.texts())
                if all(t in ui_text for t in texts):
                    return
                time.sleep(0.5)

            last_error = f"pywinauto 未找到文本，UI内容：\n{ui_text}"

        except Exception as e:
            last_error = f"pywinauto 异常：{e}"

    # ---------- 2️⃣ OCR 兜底 ----------
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    img = pyautogui.screenshot(region=ocr_region)
    ocr_text = pytesseract.image_to_string(
        img,
        lang="eng",
        config="--psm 6"
    )

    if all(t in ocr_text for t in texts):
        return

    # ---------- ❌ 断言失败 ----------
    raise AssertionError(
        "\n".join([
            "❌ 页面文本校验失败",
            f"期望包含: {texts}",
            "",
            "【OCR 识别内容】",
            ocr_text.strip(),
            "",
            "【pywinauto 信息】",
            last_error
        ])
    )

assert_page_contains(
    texts=["01 00:30:00", "02-01"],
    exe_path=r"C:\Users\CongsongLiu\Acuview2\Acuview 2.exe",
    window_title_re="Acuview",
    ocr_region=(300, 200, 1200, 600),
    tesseract_path=r"C:\Program Files\Tesseract-OCR\tesseract.exe",
)
