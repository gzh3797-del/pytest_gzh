# -*- coding: utf-8 -*-
"""
自动精度测试标签页（tkinter）

复用控源页已连接的 XL9600 + fast_accuracy_test 框架。需要宿主 app 提供：
  app.root           主窗口（用于 after 切回主线程）
  app.dev            已连接的 XL9600（None=未连接）
  app.cfg            控源页“参数配置”的 {label: StringVar}
  app._emit(msg,tag) 写共享日志
  app.set_busy(bool) 设/清“命令在途”标志（屏蔽手动命令）
"""

import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog

import accuracy_engine as eng
from fast_accuracy_test.core.config import load_config
from fast_accuracy_test.core.excel_input import Case
from xl9600 import SourceParams, to_float


def _conv(kind, raw):
    """按字段类型把界面字符串转成值；optfloat 空串 -> None。"""
    s = (raw or "").strip()
    if kind == "str":
        return s
    if kind == "optfloat":
        return None if s == "" else float(s)
    if kind == "int":
        return int(float(s))
    return float(s)


class AccuracyTab:
    def __init__(self, parent, app):
        self.app = app
        self.cases: list = []
        self._stop = threading.Event()
        self._running = False

        self._build_meter(parent)
        self._build_cases(parent)
        self._build_point(parent)
        self._build_control(parent)
        self._build_results(parent)

        self._load_config_into_ui()
        self._load_cases_initial()

    # ------------------------------------------------------------------ #
    # 读表设置
    # ------------------------------------------------------------------ #
    def _build_meter(self, parent):
        f = ttk.LabelFrame(parent, text="被检表读表设置 (Modbus)")
        f.pack(fill="x", padx=8, pady=4)
        self.m = {}

        ttk.Label(f, text="连接方式").grid(row=0, column=0, sticky="e", padx=4, pady=3)
        self.m["conn_mode"] = tk.StringVar(value="rtu")
        cb = ttk.Combobox(f, textvariable=self.m["conn_mode"], values=["rtu", "tcp"],
                          width=6, state="readonly")
        cb.grid(row=0, column=1, padx=4)
        cb.bind("<<ComboboxSelected>>", lambda e: self._sync_conn_mode())

        ttk.Label(f, text="型号").grid(row=0, column=2, sticky="e", padx=4)
        self.m["device_model"] = tk.StringVar(value="320")
        ttk.Combobox(f, textvariable=self.m["device_model"], values=["320", "300", "260"],
                     width=6, state="readonly").grid(row=0, column=3, padx=4)

        ttk.Label(f, text="字序").grid(row=0, column=4, sticky="e", padx=4)
        self.m["word_order"] = tk.StringVar(value="big")
        ttk.Combobox(f, textvariable=self.m["word_order"], values=["big", "little"],
                     width=6, state="readonly").grid(row=0, column=5, padx=4)

        ttk.Label(f, text="稳定等待(s)").grid(row=0, column=6, sticky="e", padx=4)
        self.m["settle_s"] = tk.StringVar(value="5")
        ttk.Entry(f, textvariable=self.m["settle_s"], width=5).grid(row=0, column=7, padx=4)
        ttk.Label(f, text="读重试").grid(row=0, column=8, sticky="e", padx=4)
        self.m["read_retries"] = tk.StringVar(value="3")
        ttk.Entry(f, textvariable=self.m["read_retries"], width=4).grid(row=0, column=9, padx=4)

        # 串口段
        self.rtu_frame = ttk.Frame(f)
        self.rtu_frame.grid(row=1, column=0, columnspan=10, sticky="w", pady=2)
        ttk.Label(self.rtu_frame, text="串口").grid(row=0, column=0, padx=4)
        self.m["rtu_port"] = tk.StringVar(value="COM4")
        ttk.Entry(self.rtu_frame, textvariable=self.m["rtu_port"], width=8).grid(row=0, column=1)
        ttk.Label(self.rtu_frame, text="波特率").grid(row=0, column=2, padx=4)
        self.m["rtu_baud"] = tk.StringVar(value="19200")
        ttk.Entry(self.rtu_frame, textvariable=self.m["rtu_baud"], width=8).grid(row=0, column=3)
        ttk.Label(self.rtu_frame, text="校验").grid(row=0, column=4, padx=4)
        self.m["rtu_parity"] = tk.StringVar(value="N")
        ttk.Combobox(self.rtu_frame, textvariable=self.m["rtu_parity"], values=["N", "E", "O"],
                     width=3, state="readonly").grid(row=0, column=5)
        ttk.Label(self.rtu_frame, text="从站号").grid(row=0, column=6, padx=4)
        self.m["rtu_slave"] = tk.StringVar(value="1")
        ttk.Entry(self.rtu_frame, textvariable=self.m["rtu_slave"], width=5).grid(row=0, column=7)

        # 网口段
        self.tcp_frame = ttk.Frame(f)
        self.tcp_frame.grid(row=2, column=0, columnspan=10, sticky="w", pady=2)
        ttk.Label(self.tcp_frame, text="IP").grid(row=0, column=0, padx=4)
        self.m["tcp_ip"] = tk.StringVar(value="192.168.1.254")
        ttk.Entry(self.tcp_frame, textvariable=self.m["tcp_ip"], width=16).grid(row=0, column=1)
        ttk.Label(self.tcp_frame, text="端口").grid(row=0, column=2, padx=4)
        self.m["tcp_port"] = tk.StringVar(value="502")
        ttk.Entry(self.tcp_frame, textvariable=self.m["tcp_port"], width=8).grid(row=0, column=3)
        ttk.Label(self.tcp_frame, text="从站号").grid(row=0, column=4, padx=4)
        self.m["tcp_slave"] = tk.StringVar(value="1")
        ttk.Entry(self.tcp_frame, textvariable=self.m["tcp_slave"], width=5).grid(row=0, column=5)

        ttk.Button(f, text="保存读表设置到 config.json", command=self._save_config).grid(
            row=3, column=0, columnspan=4, sticky="w", padx=4, pady=4)

    def _sync_conn_mode(self):
        rtu = self.m["conn_mode"].get() == "rtu"
        for child in self.rtu_frame.winfo_children():
            child.configure(state=("normal" if rtu else "disabled"))
        for child in self.tcp_frame.winfo_children():
            try:
                child.configure(state=("disabled" if rtu else "normal"))
            except tk.TclError:
                pass

    # ------------------------------------------------------------------ #
    # 用例表（可编辑）
    # ------------------------------------------------------------------ #
    def _build_cases(self, parent):
        f = ttk.LabelFrame(parent, text="测试用例（来自 input.xlsx，可双击单元格修改）")
        f.pack(fill="both", expand=True, padx=8, pady=4)

        bar = ttk.Frame(f)
        bar.pack(fill="x")
        ttk.Button(bar, text="加载 xlsx…", command=self._pick_xlsx).pack(side="left", padx=2)
        ttk.Button(bar, text="重新加载", command=self._load_cases_initial).pack(side="left", padx=2)
        ttk.Button(bar, text="保存回 xlsx", command=self._save_xlsx).pack(side="left", padx=2)
        ttk.Button(bar, text="导出 xlsx…", command=self._export_xlsx).pack(side="left", padx=2)
        ttk.Button(bar, text="加一行", command=self._add_row).pack(side="left", padx=2)
        ttk.Button(bar, text="删选中行", command=self._del_row).pack(side="left", padx=2)
        self.xlsx_lbl = ttk.Label(bar, text="", foreground="gray")
        self.xlsx_lbl.pack(side="left", padx=8)

        cols = [fld for fld, _, _ in eng.CASE_FIELDS]
        self.tree = ttk.Treeview(f, columns=cols, show="headings", height=7)
        for fld, title, _ in eng.CASE_FIELDS:
            self.tree.heading(fld, text=title)
            self.tree.column(fld, width=72, anchor="center")
        self.tree.column("test_case", width=90)
        ysb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=ysb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._begin_edit)

    def _render_cases(self):
        self.tree.delete(*self.tree.get_children())
        for c in self.cases:
            self.tree.insert("", "end", values=self._case_to_row(c))

    @staticmethod
    def _case_to_row(c):
        out = []
        for fld, _, _ in eng.CASE_FIELDS:
            v = getattr(c, fld, None)
            out.append("" if v is None else v)
        return out

    def _begin_edit(self, event):
        if self._running:
            return
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col:
            return
        cidx = int(col[1:]) - 1
        fld, _, kind = eng.CASE_FIELDS[cidx]
        x, y, w, h = self.tree.bbox(item, col)
        cur = self.tree.set(item, fld)
        ent = ttk.Entry(self.tree)
        ent.place(x=x, y=y, width=w, height=h)
        ent.insert(0, str(cur))
        ent.focus_set()

        def commit(_=None):
            raw = ent.get()
            ent.destroy()
            ridx = self.tree.index(item)
            try:
                setattr(self.cases[ridx], fld, _conv(kind, raw))
            except ValueError:
                self.app._emit(f"[精度] 列「{fld}」数值无效: {raw!r}", "err")
                return
            self.tree.set(item, fld, "" if raw.strip() == "" and kind == "optfloat" else raw)

        ent.bind("<Return>", commit)
        ent.bind("<FocusOut>", commit)
        ent.bind("<Escape>", lambda e: ent.destroy())

    def _add_row(self):
        self.cases.append(Case(test_case=f"case_{len(self.cases)+1}", voltage=0.0,
                               current_1=0.0, current_2=0.0, wait_h=0.0,
                               voltage_accuracy=0.001, current_accuracy=0.075,
                               power_accuracy=0.075, sample_cnt=20, sample_interval=0.1))
        self._render_cases()

    def _del_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        for item in sel:
            del self.cases[self.tree.index(item)]
        self._render_cases()

    def _pick_xlsx(self):
        path = filedialog.askopenfilename(title="选择用例 xlsx",
                                          filetypes=[("Excel", "*.xlsx")])
        if path:
            self._load_cases(path)

    def _load_cases_initial(self):
        try:
            self._load_cases(eng.input_xlsx_path())
        except Exception as e:  # noqa: BLE001
            self.app._emit(f"[精度] 默认用例加载失败: {e}", "err")

    def _load_cases(self, path):
        try:
            self.cases = eng.load_cases(path)
        except Exception as e:  # noqa: BLE001
            self.app._emit(f"[精度] 读用例失败: {e}", "err")
            return
        self._xlsx_path = path
        self.xlsx_lbl.config(text=f"{len(self.cases)} 条用例  ←  {path}")
        self._render_cases()

    def _save_xlsx(self):
        try:
            path = eng.save_cases_xlsx(self.cases, getattr(self, "_xlsx_path", None))
            self.app._emit(f"[精度] 用例已保存回 {path}", "ok")
        except Exception as e:  # noqa: BLE001
            self.app._emit(f"[精度] 保存 xlsx 失败（是否被 Excel 占用？）: {e}", "err")

    def _export_xlsx(self):
        if not self.cases:
            self.app._emit("[精度] 没有用例可导出", "err")
            return
        path = filedialog.asksaveasfilename(
            title="导出用例为 xlsx", defaultextension=".xlsx",
            initialfile="cases_export.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            out = eng.export_cases_xlsx(self.cases, path)
            self.app._emit(f"[精度] 已导出 {len(self.cases)} 条用例到 {out}", "ok")
        except Exception as e:  # noqa: BLE001
            self.app._emit(f"[精度] 导出失败（目标是否被占用？）: {e}", "err")

    # ------------------------------------------------------------------ #
    # 单点快测
    # ------------------------------------------------------------------ #
    def _build_point(self, parent):
        f = ttk.LabelFrame(parent, text="单点快测（控源→稳定→读表→打印误差，不出报告）")
        f.pack(fill="x", padx=8, pady=4)
        ttk.Label(f, text="电压(V)").grid(row=0, column=0, padx=4, pady=4)
        self.pt_v = tk.StringVar(value="60")
        ttk.Entry(f, textvariable=self.pt_v, width=8).grid(row=0, column=1)
        ttk.Label(f, text="电流(A)").grid(row=0, column=2, padx=4)
        self.pt_i = tk.StringVar(value="1.2")
        ttk.Entry(f, textvariable=self.pt_i, width=8).grid(row=0, column=3)
        ttk.Label(f, text="压/流/功精度(可空)").grid(row=0, column=4, padx=4)
        self.pt_acc = tk.StringVar(value="")
        ttk.Entry(f, textvariable=self.pt_acc, width=16).grid(row=0, column=5)
        self.btn_point = ttk.Button(f, text="单点快测", command=self._do_point)
        self.btn_point.grid(row=0, column=6, padx=8)

    # ------------------------------------------------------------------ #
    # 控制条 + 结果表
    # ------------------------------------------------------------------ #
    def _build_control(self, parent):
        f = ttk.Frame(parent)
        f.pack(fill="x", padx=8, pady=4)
        self.btn_start = ttk.Button(f, text="开始测试", command=self._start)
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = ttk.Button(f, text="停止", command=self._request_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        self.prog = ttk.Label(f, text="就绪")
        self.prog.pack(side="left", padx=12)
        self.btn_report = ttk.Button(f, text="打开报告", command=self._open_report, state="disabled")
        self.btn_report.pack(side="right", padx=4)
        self._report_path = ""

    def _build_results(self, parent):
        f = ttk.LabelFrame(parent, text="结果")
        f.pack(fill="both", expand=True, padx=8, pady=4)
        self.res = ttk.Treeview(f, columns=("case", "overall", "detail"),
                                show="headings", height=6)
        for cid, title, w in (("case", "用例", 100), ("overall", "总判定", 80),
                              ("detail", "明细（项=判定）", 600)):
            self.res.heading(cid, text=title)
            self.res.column(cid, width=w, anchor="w")
        self.res.tag_configure("pass", foreground="green")
        self.res.tag_configure("fail", foreground="red")
        self.res.pack(fill="both", expand=True)

    # ------------------------------------------------------------------ #
    # config.json 读写
    # ------------------------------------------------------------------ #
    def _load_config_into_ui(self):
        try:
            cfg = load_config(eng.default_config_path())
        except Exception as e:  # noqa: BLE001
            self.app._emit(f"[精度] 读 config.json 失败: {e}", "err")
            return
        self.m["conn_mode"].set(cfg.get("conn_mode", "rtu"))
        self.m["device_model"].set(cfg.get("device_model", "320"))
        self.m["word_order"].set(cfg.get("word_order", "big"))
        self.m["settle_s"].set(str(cfg.get("settle_s", 5)))
        self.m["read_retries"].set(str(cfg.get("read_retries", 3)))
        rtu, tcp = cfg.get("rtu", {}), cfg.get("tcp", {})
        self.m["rtu_port"].set(rtu.get("port", "COM4"))
        self.m["rtu_baud"].set(str(rtu.get("baudrate", 19200)))
        self.m["rtu_parity"].set(rtu.get("parity", "N"))
        self.m["rtu_slave"].set(str(rtu.get("slaveid", 1)))
        self.m["tcp_ip"].set(tcp.get("ip", "192.168.1.254"))
        self.m["tcp_port"].set(str(tcp.get("port", 502)))
        self.m["tcp_slave"].set(str(tcp.get("slaveid", 1)))
        self._sync_conn_mode()

    def _meter_dict(self):
        return {
            "conn_mode": self.m["conn_mode"].get(),
            "device_model": self.m["device_model"].get(),
            "word_order": self.m["word_order"].get(),
            "settle_s": float(self.m["settle_s"].get()),
            "read_retries": int(self.m["read_retries"].get()),
            "rtu": {"port": self.m["rtu_port"].get(), "baudrate": int(self.m["rtu_baud"].get()),
                    "parity": self.m["rtu_parity"].get(), "slaveid": int(self.m["rtu_slave"].get())},
            "tcp": {"ip": self.m["tcp_ip"].get(), "port": int(self.m["tcp_port"].get()),
                    "slaveid": int(self.m["tcp_slave"].get())},
        }

    def _save_config(self):
        path = eng.default_config_path()
        try:
            with open(path, "r", encoding="utf-8") as fp:
                raw = json.load(fp)
            m = self._meter_dict()
            raw.update({k: m[k] for k in ("conn_mode", "device_model", "word_order",
                                          "settle_s", "read_retries", "rtu", "tcp")})
            with open(path, "w", encoding="utf-8") as fp:
                json.dump(raw, fp, ensure_ascii=False, indent=2)
            self.app._emit(f"[精度] 读表设置已保存到 {path}", "ok")
        except Exception as e:  # noqa: BLE001
            self.app._emit(f"[精度] 保存 config.json 失败: {e}", "err")

    # ------------------------------------------------------------------ #
    # 控源页参数（额定基准 + 参数配置帧）
    # ------------------------------------------------------------------ #
    def _rated_and_params(self):
        c = self.app.cfg
        rated_v = to_float(c["额定电压"].get())
        rated_i = to_float(c["标定电流"].get())
        params = SourceParams(
            电流接入方式=c["电流接入方式"].get(), 供电方式=c["供电方式"].get(),
            额定电压=c["额定电压"].get(), 标定电流=c["标定电流"].get(),
            分流器额定=c["分流器额定"].get(), 被检表阻抗=c["被检表阻抗"].get(),
            脉冲常数=c["脉冲常数"].get(), 校验圈数=c["校验圈数"].get(),
            校验秒数=c["校验秒数"].get())
        return rated_v, rated_i, params

    # ------------------------------------------------------------------ #
    # 跑测试
    # ------------------------------------------------------------------ #
    def _precheck(self):
        if self.app.dev is None:
            self.app._emit("[精度] 请先到「手动控源」页连接 XL9600", "err")
            return False
        if self._running or self.app._busy:
            self.app._emit("[精度] 正忙，请稍候", "err")
            return False
        return True

    def _set_running(self, running):
        self._running = running
        self.app.set_busy(running)
        self.btn_start.config(state="disabled" if running else "normal")
        self.btn_point.config(state="disabled" if running else "normal")
        self.btn_stop.config(state="normal" if running else "disabled")

    def _start(self):
        if not self._precheck():
            return
        if not self.cases:
            self.app._emit("[精度] 没有用例，请先加载/添加", "err")
            return
        try:
            rated_v, rated_i, params = self._rated_and_params()
            meter = self._meter_dict()
        except (ValueError, KeyError) as e:
            self.app._emit(f"[精度] 参数无效: {e}", "err")
            return
        self.res.delete(*self.res.get_children())
        self._stop.clear()
        self._set_running(True)
        self.app._emit(f"[精度] 开始测试，共 {len(self.cases)} 条用例", "tx")

        cases = list(self.cases)
        dev = self.app.dev

        def progress(kind, payload):
            self.app.root.after(0, lambda: self._on_progress(kind, payload))

        def worker():
            try:
                eng.run_cases(dev, rated_v, rated_i, meter, cases,
                              source_params=params, progress=progress, stop_event=self._stop)
            except Exception as e:  # noqa: BLE001
                progress("warn", f"测试中止: {e}")
            finally:
                self.app.root.after(0, lambda: self._set_running(False))

        threading.Thread(target=worker, daemon=True).start()

    def _do_point(self):
        if not self._precheck():
            return
        try:
            rated_v, rated_i, params = self._rated_and_params()
            meter = self._meter_dict()
            v, i = float(self.pt_v.get()), float(self.pt_i.get())
            acc = [float(x) for x in self.pt_acc.get().split() if x.strip()]
        except (ValueError, KeyError) as e:
            self.app._emit(f"[精度] 参数无效: {e}", "err")
            return
        kw = {}
        if len(acc) == 3:
            kw = dict(vacc=acc[0], iacc=acc[1], pacc=acc[2])
        self._stop.clear()
        self._set_running(True)
        self.app._emit(f"[精度] 单点快测 {v}V / {i}A …", "tx")
        dev = self.app.dev

        def worker():
            try:
                res = eng.run_point(dev, rated_v, rated_i, meter, v, i,
                                    source_params=params, stop_event=self._stop, **kw)
                self.app.root.after(0, lambda: self._on_progress("case_done", res))
            except Exception as e:  # noqa: BLE001
                self.app.root.after(0, lambda: self.app._emit(f"[精度] 单点快测失败: {e}", "err"))
            finally:
                self.app.root.after(0, lambda: self._set_running(False))

        threading.Thread(target=worker, daemon=True).start()

    def _request_stop(self):
        self._stop.set()
        self.app._emit("[精度] 已请求停止，当前用例结束后停止…", "tx")

    # ------------------------------------------------------------------ #
    # 进度/结果回调（主线程）
    # ------------------------------------------------------------------ #
    def _on_progress(self, kind, payload):
        if kind == "total":
            self._total = payload
            self._done = 0
            self.prog.config(text=f"0/{payload}")
        elif kind == "case_start":
            self.prog.config(text=f"{getattr(self,'_done',0)}/{getattr(self,'_total','?')}  正在测 {payload}")
        elif kind == "case_done":
            self._done = getattr(self, "_done", 0) + 1
            self.prog.config(text=f"{self._done}/{getattr(self,'_total','?')}")
            self._add_result(payload)
        elif kind == "stopped":
            self.prog.config(text="已停止")
            self.app._emit("[精度] 测试已停止", "err")
        elif kind == "report":
            self._report_path = payload
            self.btn_report.config(state="normal")
            self.prog.config(text=self.prog.cget("text") + "  完成")
            self.app._emit(f"[精度] 报告已生成: {payload}", "ok")
        elif kind == "warn":
            self.app._emit(f"[精度] {payload}", "err")

    def _add_result(self, res):
        detail = "  ".join(f"{m.label}={m.result}" for m in res.metrics)
        tag = "pass" if res.overall == "Passed" else "fail"
        self.res.insert("", "end", values=(res.case.test_case, res.overall, detail), tags=(tag,))
        # 明细同时进共享日志
        for m in res.metrics:
            if m.result == "N/A":
                continue
            self.app._emit(
                f"   {res.case.test_case} {m.label}: 真值={_fmt(m.ref)} 平均={_fmt(m.avg)} "
                f"平均误差={_fmt(m.err_avg)}% 最大误差={_fmt(m.err_worst)}% -> {m.result}",
                "ok" if m.result == "Passed" else "err")

    def _open_report(self):
        if not self._report_path:
            return
        try:
            import os
            os.startfile(self._report_path)  # Windows
        except Exception as e:  # noqa: BLE001
            self.app._emit(f"[精度] 打开报告失败: {e}", "err")


def _fmt(x):
    return "-" if x is None else ("%.5g" % x)
