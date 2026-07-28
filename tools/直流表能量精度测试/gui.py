# -*- coding: utf-8 -*-
"""
XL-9600 直流电能表检定装置 控源界面 (tkinter)

依赖：仅 Python 标准库 + 同目录 xl9600.py
运行：python gui.py
"""

import queue
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext

from xl9600 import (
    XL9600, SourceParams, OutputPoint,
    XL9600Error, XL9600Timeout, to_float,
)


class XL9600GUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("XL-9600 直流电能表检定装置 控源")
        root.geometry("900x960")

        self.dev: XL9600 | None = None
        self.log_queue: "queue.Queue[tuple[str, str]]" = queue.Queue()
        self._busy = False  # 命令在途标志，防止连点叠加

        # 底部共享日志（常驻，两个标签页共用），先 pack 到底部
        self._build_log()

        # 两个标签页：手动控源 / 自动精度测试
        nb = ttk.Notebook(root)
        nb.pack(side="top", fill="both", expand=True)
        self.tab1 = ttk.Frame(nb)
        self.tab2 = ttk.Frame(nb)
        nb.add(self.tab1, text="手动控源")
        nb.add(self.tab2, text="自动精度测试")

        # Tab1：原控源功能
        self._build_conn()
        self._build_config()
        self._build_output()
        self._build_error()
        self._build_actions()

        # Tab2：精度测试（复用本页连接 + 共享日志）
        from accuracy_tab import AccuracyTab
        self.acc = AccuracyTab(self.tab2, self)

        self._set_connected(False)
        self.root.after(100, self._drain_log)
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    def set_busy(self, busy: bool) -> None:
        """供精度测试页设/清“命令在途”标志，屏蔽手动命令叠加。"""
        self._busy = busy

    # ------------------------------------------------------------------ #
    # 控件布局
    # ------------------------------------------------------------------ #
    def _build_conn(self) -> None:
        f = ttk.LabelFrame(self.tab1, text="连接")
        f.pack(fill="x", padx=8, pady=4)
        ttk.Label(f, text="IP").grid(row=0, column=0, padx=4, pady=4)
        self.ip = tk.StringVar(value="192.168.1.105")
        ttk.Entry(f, textvariable=self.ip, width=16).grid(row=0, column=1)
        ttk.Label(f, text="端口").grid(row=0, column=2, padx=4)
        self.port = tk.StringVar(value="24433")
        ttk.Entry(f, textvariable=self.port, width=8).grid(row=0, column=3)
        ttk.Label(f, text="超时(s)").grid(row=0, column=4, padx=4)
        self.timeout = tk.StringVar(value="5")
        ttk.Entry(f, textvariable=self.timeout, width=6).grid(row=0, column=5)
        self.btn_conn = ttk.Button(f, text="连接", command=self._toggle_conn)
        self.btn_conn.grid(row=0, column=6, padx=8)
        self.lbl_state = ttk.Label(f, text="● 未连接", foreground="gray")
        self.lbl_state.grid(row=0, column=7, padx=4)

    def _build_config(self) -> None:
        f = ttk.LabelFrame(self.tab1, text="参数配置")
        f.pack(fill="x", padx=8, pady=4)
        self.cfg: dict[str, tk.StringVar] = {}

        def row(r, c, label, var, values=None, width=12):
            ttk.Label(f, text=label).grid(row=r, column=c * 2, sticky="e", padx=4, pady=3)
            sv = tk.StringVar(value=var)
            self.cfg[label] = sv
            if values:
                ttk.Combobox(f, textvariable=sv, values=values, width=width - 2,
                             state="readonly").grid(row=r, column=c * 2 + 1, padx=4)
            else:
                ttk.Entry(f, textvariable=sv, width=width).grid(row=r, column=c * 2 + 1, padx=4)

        row(0, 0, "电流接入方式", "间接接入式", ["直接接入式", "间接接入式"])
        row(0, 1, "供电方式", "电源供电", ["电源供电", "线路供电"])
        row(1, 0, "额定电压", "100V")
        row(1, 1, "标定电流", "100A")
        row(2, 0, "分流器额定", "75mV")
        row(2, 1, "被检表阻抗", "1000Ω")
        row(3, 0, "脉冲常数", "1000")
        row(3, 1, "校验圈数", "自动")
        row(4, 0, "校验秒数", "1")
        ttk.Button(f, text="下发参数配置", command=self._do_config).grid(
            row=4, column=3, padx=8, pady=4, sticky="e")

    def _build_output(self) -> None:
        f = ttk.LabelFrame(self.tab1, text="源输出")
        f.pack(fill="x", padx=8, pady=4)
        self.out: dict[str, tk.StringVar] = {}

        def row(r, c, label, var, values=None, width=10):
            ttk.Label(f, text=label).grid(row=r, column=c * 2, sticky="e", padx=4, pady=3)
            sv = tk.StringVar(value=var)
            self.out[label] = sv
            if values:
                ttk.Combobox(f, textvariable=sv, values=values, width=width - 2,
                             state="readonly").grid(row=r, column=c * 2 + 1, padx=4)
            else:
                ttk.Entry(f, textvariable=sv, width=width).grid(row=r, column=c * 2 + 1, padx=4)

        row(0, 0, "输出电压(V)", "100")
        row(0, 1, "输出电流(A)", "100")
        ttk.Label(f, text="(填实际值，程序按额定电压/标定电流自动换算成%)",
                  foreground="gray").grid(row=0, column=4, columnspan=2, sticky="w", padx=4)
        row(1, 0, "电压纹波比例", "0%")
        row(1, 1, "电流纹波比例", "0%")
        row(2, 0, "电压纹波相位", "0度")
        row(2, 1, "电流纹波相位", "0度")
        row(3, 0, "纹波频率", "300Hz")
        row(3, 1, "电能方向", "正向", ["正向", "反向"])
        self.btn_out = ttk.Button(f, text="源输出", command=self._do_output)
        self.btn_out.grid(row=4, column=1, padx=4, pady=4, sticky="w")
        self.btn_stop = ttk.Button(f, text="源停止", command=self._do_stop)
        self.btn_stop.grid(row=4, column=3, padx=4, pady=4, sticky="w")

    def _build_error(self) -> None:
        f = ttk.LabelFrame(self.tab1, text="误差读取")
        f.pack(fill="x", padx=8, pady=4)
        ttk.Label(f, text="统计次数").grid(row=0, column=0, padx=4, pady=4)
        self.cnt = tk.StringVar(value="5")
        ttk.Entry(f, textvariable=self.cnt, width=8).grid(row=0, column=1)
        self.btn_err = ttk.Button(f, text="读电能误差", command=self._do_read_error)
        self.btn_err.grid(row=0, column=2, padx=6)
        self.btn_clk = ttk.Button(f, text="读日计时误差", command=self._do_read_clock)
        self.btn_clk.grid(row=0, column=3, padx=6)
        ttk.Label(f, text="均值").grid(row=0, column=4, padx=4)
        self.mean = tk.StringVar(value="-")
        ttk.Entry(f, textvariable=self.mean, width=14, state="readonly").grid(row=0, column=5)

    def _build_actions(self) -> None:
        f = ttk.Frame(self.tab1)
        f.pack(fill="x", padx=8, pady=4)
        self.btn_off = ttk.Button(f, text="供电关闭（关闭所有输出）", command=self._do_power_off)
        self.btn_off.pack(side="left", padx=4)
        ttk.Button(f, text="清空日志", command=self._clear_log).pack(side="right", padx=4)

    def _build_log(self) -> None:
        f = ttk.LabelFrame(self.root, text="通信日志（共享）")
        f.pack(side="bottom", fill="x", padx=8, pady=4)
        self.log = scrolledtext.ScrolledText(f, height=10, state="disabled",
                                             font=("Consolas", 9))
        self.log.pack(fill="both", expand=True)
        self.log.tag_config("err", foreground="red")
        self.log.tag_config("ok", foreground="green")
        self.log.tag_config("tx", foreground="blue")

    # ------------------------------------------------------------------ #
    # 日志（线程安全）
    # ------------------------------------------------------------------ #
    def _emit(self, msg: str, tag: str = "") -> None:
        self.log_queue.put((msg, tag))

    def _drain_log(self) -> None:
        try:
            while True:
                msg, tag = self.log_queue.get_nowait()
                self.log.configure(state="normal")
                self.log.insert("end", msg + "\n", tag)
                self.log.see("end")
                self.log.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log)

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------ #
    # 连接
    # ------------------------------------------------------------------ #
    def _toggle_conn(self) -> None:
        if self.dev is None:
            try:
                dev = XL9600(self.ip.get().strip(), int(self.port.get()),
                             timeout=float(self.timeout.get()))
                dev.open()
            except Exception as e:  # noqa: BLE001
                self._emit(f"[连接失败] {e}", "err")
                return
            # UDP 无连接：必须发探测命令并收到应答，才确认设备真的在线。
            # 探测放后台线程，避免设备不在线时界面卡住 timeout 秒。
            ip, port = self.ip.get(), self.port.get()
            self._emit(f"-> 探测设备 {ip}:{port} ...", "tx")
            self.btn_conn.config(state="disabled")

            def probe():
                try:
                    online = dev.ping()
                except Exception as e:  # noqa: BLE001
                    self.root.after(0, lambda: self._finish_connect(dev, None, str(e)))
                    return
                self.root.after(0, lambda: self._finish_connect(dev, online, None))

            threading.Thread(target=probe, daemon=True).start()
        else:
            self.dev.close()
            self.dev = None
            self._emit("[已断开]", "")
            self._set_connected(False)

    def _finish_connect(self, dev, online, error) -> None:
        """探测结果回到主线程处理。"""
        self.btn_conn.config(state="normal")
        if error is not None:
            dev.close()
            self._emit(f"[连接失败] 探测出错: {error}", "err")
            return
        if not online:
            dev.close()
            self._emit("[连接失败] 设备无应答，请检查 IP / 端口 / 网络是否连通", "err")
            return
        self.dev = dev
        self._emit(f"[已连接] {self.ip.get()}:{self.port.get()} 设备已响应", "ok")
        self._set_connected(True)

    def _set_connected(self, connected: bool) -> None:
        self.btn_conn.config(text="断开" if connected else "连接")
        self.lbl_state.config(text="● 已连接" if connected else "● 未连接",
                              foreground="green" if connected else "gray")
        state = "normal" if connected else "disabled"
        for b in (self.btn_out, self.btn_stop, self.btn_err, self.btn_clk, self.btn_off):
            b.config(state=state)

    # ------------------------------------------------------------------ #
    # 命令执行（后台线程，避免界面卡死）
    # ------------------------------------------------------------------ #
    def _run_async(self, func, name: str) -> None:
        if self.dev is None:
            self._emit("[未连接] 请先连接设备", "err")
            return
        if self._busy:
            self._emit(f"[忙] 上一条命令未完成，已忽略「{name}」", "err")
            return
        self._busy = True

        def worker():
            try:
                self._emit(f"-> 发送 {name}", "tx")
                func(self.dev)
            except XL9600Timeout as e:
                self._emit(f"[超时] {e}", "err")
            except XL9600Error as e:
                self._emit(f"[设备错误] {e}", "err")
            except Exception as e:  # noqa: BLE001
                self._emit(f"[错误] {e}", "err")
            finally:
                self.root.after(0, self._mark_idle)

        threading.Thread(target=worker, daemon=True).start()

    def _mark_idle(self) -> None:
        self._busy = False

    def _do_config(self) -> None:
        c = self.cfg
        params = SourceParams(
            电流接入方式=c["电流接入方式"].get(),
            供电方式=c["供电方式"].get(),
            额定电压=c["额定电压"].get(),
            标定电流=c["标定电流"].get(),
            分流器额定=c["分流器额定"].get(),
            被检表阻抗=c["被检表阻抗"].get(),
            脉冲常数=c["脉冲常数"].get(),
            校验圈数=c["校验圈数"].get(),
            校验秒数=c["校验秒数"].get(),
        )

        def f(dev):
            dev.config(params)
            self._emit("[OK] 参数配置已下发", "ok")
        self._run_async(f, "参数配置")

    def _do_output(self) -> None:
        o = self.out
        # 把实际电压/电流换算成协议要求的百分比检定点
        try:
            rated_v = to_float(self.cfg["额定电压"].get())
            rated_i = to_float(self.cfg["标定电流"].get())
            act_v = to_float(o["输出电压(V)"].get())
            act_i = to_float(o["输出电流(A)"].get())
        except ValueError as e:
            self._emit(f"[错误] 电压/电流数值无法解析: {e}", "err")
            return
        if rated_v <= 0 or rated_i <= 0:
            self._emit("[错误] 额定电压/标定电流必须大于 0（请先在参数配置里填好）", "err")
            return

        pct_v = act_v / rated_v * 100.0
        pct_i = act_i / rated_i * 100.0
        self._emit(f"   换算: 电压 {act_v:g}V / 额定 {rated_v:g}V = {pct_v:g}%  |  "
                   f"电流 {act_i:g}A / 标定 {rated_i:g}A = {pct_i:g}%", "tx")

        point = OutputPoint(
            电压检定点=f"{pct_v:g}%",
            电流检定点=f"{pct_i:g}%",
            电压纹波比例=o["电压纹波比例"].get(),
            电流纹波比例=o["电流纹波比例"].get(),
            电压纹波相位=o["电压纹波相位"].get(),
            电流纹波相位=o["电流纹波相位"].get(),
            纹波频率=o["纹波频率"].get(),
            电能方向=o["电能方向"].get(),
        )

        def f(dev):
            r = dev.source_output(point)
            self._emit(f"[OK] 源输出  电压总值={r.get('电压总值')} "
                       f"电流总值={r.get('电流总值')} 功率总值={r.get('功率总值')}", "ok")
        self._run_async(f, "源输出")

    def _do_stop(self) -> None:
        def f(dev):
            dev.source_stop()
            self._emit("[OK] 源已停止", "ok")
        self._run_async(f, "源停止")

    def _do_read_error(self) -> None:
        try:
            n = int(self.cnt.get())
        except ValueError:
            self._emit("[错误] 统计次数必须是整数", "err")
            return

        def f(dev):
            res = dev.read_error(统计次数=n)
            self.mean.set(str(res.均值))
            if res.均值 != res.均值:  # nan：没解析到均值，回显原始报文便于排查
                self._emit(f"[注意] 未解析到均值，设备原始回复: {res.raw.get('_raw')!r}", "err")
            else:
                self._emit(f"[OK] 电能误差 均值={res.均值} 原始值={res.原始值}", "ok")
        self._run_async(f, "误差读取")

    def _do_read_clock(self) -> None:
        try:
            n = int(self.cnt.get())
        except ValueError:
            self._emit("[错误] 统计次数必须是整数", "err")
            return

        def f(dev):
            res = dev.read_clock_error(统计次数=n)
            self.mean.set(f"{res.均值} s")
            self._emit(f"[OK] 日计时误差 均值={res.均值}s 原始值={res.原始值}", "ok")
        self._run_async(f, "日计时误差读取")

    def _do_power_off(self) -> None:
        def f(dev):
            dev.power_off()
            self._emit("[OK] 供电已关闭", "ok")
        self._run_async(f, "供电关闭")

    # ------------------------------------------------------------------ #
    def _on_close(self) -> None:
        if self.dev is not None:
            try:
                self.dev.close()
            except Exception:  # noqa: BLE001
                pass
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")  # Windows 原生外观
    except tk.TclError:
        pass
    XL9600GUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
