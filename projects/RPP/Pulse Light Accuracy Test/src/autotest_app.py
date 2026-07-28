# src/autotest_app.py
#
# 自动测试 GUI (config-driven, multi-wiring) — 所有设备/寄存器设置来自 config.yaml。
# GUI 本身只是运行控制台：加载 config、选 xlsx + 多张工作表、点击开始。
#
# SAFETY NOTE: 该窗口会按 xlsx 表格设置源的真实电压/电流输出；
# 启动是用户的明确操作，请确保接线正确后再点击"开始测试"。

import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from src.config_loader import (
    load_config, build_source, build_counter, build_meter,
    pulse_cfg, wiring_cfg, get_test_params,
)
from src.table_io import list_testpoint_sheets, read_test_points, write_results
from src.autotest import run_test


class AutoTestApp:
    def __init__(self, root):
        self.root = root
        root.title("能表自动测试")
        self._stop = False
        self._abort = False
        self.cfg = None

        # threading.Event for worker/main-thread pause handshake
        self._resume = threading.Event()
        self._pause_result: dict = {}

        self._build_config(root)
        self._build_test(root)
        self._build_control(root)

        # Try to auto-load config.yaml on startup (must not raise if absent)
        try:
            self._load_config_path("config.yaml")
        except Exception:
            pass  # absent or broken — status label shows red, self.cfg stays None

    # ------------------------------------------------------------------ #
    #  Section 1 — 配置
    # ------------------------------------------------------------------ #

    def _build_config(self, root):
        f = ttk.LabelFrame(root, text="配置")
        f.pack(fill="x", padx=8, pady=4)

        row = ttk.Frame(f)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="config.yaml:").pack(side="left", padx=4)
        self.e_cfg = ttk.Entry(row, width=48)
        self.e_cfg.insert(0, "config.yaml")
        self.e_cfg.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览", command=self._on_browse_config).pack(side="left", padx=4)
        ttk.Button(row, text="重新加载", command=self._on_reload_config).pack(side="left", padx=2)
        self.lbl_cfg_status = ttk.Label(row, text="未加载", foreground="red")
        self.lbl_cfg_status.pack(side="left", padx=8)

    def _load_config_path(self, path: str):
        """Load config from *path*; update status label; raises ValueError on failure."""
        self.cfg = load_config(path)
        self.lbl_cfg_status.config(
            text=f"已加载 {path}", foreground="green"
        )

    def _on_browse_config(self):
        p = filedialog.askopenfilename(
            title="选择 config.yaml",
            filetypes=[("YAML 文件", "*.yaml *.yml"), ("所有文件", "*.*")],
        )
        if not p:
            return
        self.e_cfg.delete(0, "end")
        self.e_cfg.insert(0, p)
        try:
            self._load_config_path(p)
        except Exception as e:
            messagebox.showerror("加载失败", str(e))
            self.lbl_cfg_status.config(text="加载失败", foreground="red")
            self.cfg = None

    def _on_reload_config(self):
        p = self.e_cfg.get().strip()
        try:
            self._load_config_path(p)
        except Exception as e:
            self.lbl_cfg_status.config(
                text="未找到" if "无法读取" in str(e) else "加载失败",
                foreground="red",
            )
            messagebox.showerror("加载失败", str(e))
            self.cfg = None

    # ------------------------------------------------------------------ #
    #  Section 2 — 测试
    # ------------------------------------------------------------------ #

    def _build_test(self, root):
        f = ttk.LabelFrame(root, text="测试")
        f.pack(fill="x", padx=8, pady=4)

        # xlsx row
        row1 = ttk.Frame(f)
        row1.pack(fill="x", pady=2)
        ttk.Label(row1, text="xlsx 文件:").pack(side="left", padx=4)
        self.e_xlsx = ttk.Entry(row1, width=48)
        self.e_xlsx.pack(side="left", fill="x", expand=True)
        ttk.Button(row1, text="浏览…", command=self._on_browse_xlsx).pack(side="left", padx=4)

        # sheets listbox + CT类型
        row2 = ttk.Frame(f)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="工作表(可多选):").pack(side="left", padx=4)
        lst_frame = ttk.Frame(row2)
        lst_frame.pack(side="left", padx=4)
        self.lst_sheets = tk.Listbox(
            lst_frame, selectmode="extended", height=5, width=40,
            exportselection=False,
        )
        sb_lst = ttk.Scrollbar(lst_frame, orient="vertical",
                               command=self.lst_sheets.yview)
        self.lst_sheets.configure(yscrollcommand=sb_lst.set)
        self.lst_sheets.pack(side="left")
        sb_lst.pack(side="right", fill="y")

        ttk.Label(row2, text="CT类型:").pack(side="left", padx=8)
        self.cb_ct = ttk.Combobox(row2, values=["全部", "mV", "mA"],
                                   state="readonly", width=6)
        self.cb_ct.current(0)
        self.cb_ct.pack(side="left", padx=4)

    def _on_browse_xlsx(self):
        p = filedialog.askopenfilename(
            title="选择测试点 xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not p:
            return
        self.e_xlsx.delete(0, "end")
        self.e_xlsx.insert(0, p)
        self.lst_sheets.delete(0, "end")
        try:
            sheets = list_testpoint_sheets(p)
        except Exception as e:
            messagebox.showerror("读取工作表失败", str(e))
            return
        for s in sheets:
            self.lst_sheets.insert("end", s)

    # ------------------------------------------------------------------ #
    #  Section 3 — 控制条 + 结果 Treeview
    # ------------------------------------------------------------------ #

    def _build_control(self, root):
        ctrl = ttk.Frame(root)
        ctrl.pack(fill="x", padx=8, pady=4)

        self.btn_start = ttk.Button(ctrl, text="开始测试", command=self.on_start)
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = ttk.Button(ctrl, text="停止", command=self.on_stop,
                                   state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        self.lbl_prog = ttk.Label(ctrl, text="就绪")
        self.lbl_prog.pack(side="left", padx=8)

        # 结果 Treeview — 接线 column added first to distinguish multi-wiring rows
        cols = ("接线", "行", "电压", "电流", "PF", "min(s)", "max(s)", "avg(s)", "状态")
        tree_frame = ttk.Frame(root)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=10)
        col_widths = {
            "接线": 140, "行": 40, "电压": 70, "电流": 70, "PF": 60,
            "min(s)": 90, "max(s)": 90, "avg(s)": 90, "状态": 120,
        }
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=col_widths.get(c, 80), anchor="center")
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    # ------------------------------------------------------------------ #
    #  on_start / on_stop
    # ------------------------------------------------------------------ #

    def on_start(self):
        if self.btn_start["state"] == "disabled":
            return

        # Validate: config loaded
        if self.cfg is None:
            messagebox.showwarning("未加载配置", "请先加载 config.yaml")
            return

        # Validate: xlsx chosen
        xlsx = self.e_xlsx.get().strip()
        if not xlsx:
            messagebox.showwarning("缺少文件", "请先选择 xlsx 测试点文件")
            return

        # Validate: at least one sheet selected
        sel = self.lst_sheets.curselection()
        if not sel:
            messagebox.showwarning("未选工作表", "请在工作表列表中选择至少一张工作表")
            return

        # Snapshot all main-thread state before spawning worker
        sheets = [self.lst_sheets.get(i) for i in sel]
        params = {
            "xlsx":    xlsx,
            "sheets":  sheets,
            "ct_type": self.cb_ct.get().strip(),
            "cfg":     self.cfg,  # read-only in worker
        }

        self._stop = False
        self._abort = False
        self._resume.clear()
        self._pause_result.clear()
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.lbl_prog.config(text="准备中…")
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)

        threading.Thread(target=self._run_worker, args=(params,), daemon=True).start()

    def on_stop(self):
        self._stop = True
        self.root.after(0, self.lbl_prog.config, {"text": "停止中…"})
        # Unblock any worker waiting on a pause dialog (pause_result stays empty → ok=False)
        self._resume.set()

    # ------------------------------------------------------------------ #
    #  Worker (daemon thread — MUST NOT touch Tk widgets directly)
    # ------------------------------------------------------------------ #

    def _run_worker(self, p):
        source = counter = meter = None
        try:
            cfg = p["cfg"]

            # Connect source
            self.root.after(0, self.lbl_prog.config, {"text": "连接源…"})
            source = build_source(cfg)
            source.connect()

            # Connect counter + apply fixed settings
            self.root.after(0, self.lbl_prog.config, {"text": "连接频率计…"})
            counter = build_counter(cfg)
            counter.connect()
            counter.set_counter(True)   # enable FCNT
            counter.set_mode("DC")      # 直流
            counter.set_hfr(True)       # 高频抑制打开
            counter.set_trg(1.5)        # 触发电平 1.5 V
            counter.set_type("FAST")    # 快速测量

            # Connect meter
            self.root.after(0, self.lbl_prog.config, {"text": "连接电表…"})
            meter = build_meter(cfg)
            meter.connect()

            pc = pulse_cfg(cfg)
            wc = wiring_cfg(cfg)
            tp = get_test_params(cfg)

            for sheet in p["sheets"]:
                if self._stop:
                    break

                # Look up wiring value for this sheet
                wval = wc["map"].get(sheet)
                if wval is None:
                    self.root.after(
                        0, self._insert_skip_row, sheet,
                        "未在config.wiring.map中,跳过",
                    )
                    continue

                # Write wiring register
                try:
                    meter.write_value(
                        wc["register"], wval,
                        dtype=wc.get("dtype", "uint16"),
                    )
                except Exception as exc:
                    self.root.after(
                        0, self._insert_skip_row, sheet,
                        f"写接线寄存器失败:{exc}",
                    )
                    continue

                # --- Pause: ask operator to confirm physical wiring ---
                self._resume.clear()
                self._pause_result.clear()
                self.root.after(0, self._ask_continue, sheet, wval)
                self._resume.wait()
                if not self._pause_result.get("ok") or self._stop:
                    self._stop = True
                    break

                # Status: reading test points
                self.root.after(
                    0, self.lbl_prog.config,
                    {"text": f"【{sheet}】读取测试点…"},
                )

                # Read test points
                try:
                    points = read_test_points(
                        p["xlsx"], sheet=sheet, ct_type=p["ct_type"]
                    )
                except Exception as exc:
                    self.root.after(
                        0, self._insert_skip_row, sheet,
                        f"读取测试点失败:{exc}",
                    )
                    continue

                total = len(points)

                def _make_point_start_cb(sheet_name, total_pts):
                    def cb(i, _total, _point):
                        self.root.after(
                            0, self.lbl_prog.config,
                            {"text": f"【{sheet_name}】第 {i + 1}/{total_pts} 点…"},
                        )
                    return cb

                def _make_progress_cb(sheet_name):
                    def cb(i, _total, _point, result):
                        self.root.after(0, self._on_progress, sheet_name, i, result)
                    return cb

                results = run_test(
                    points,
                    source,
                    meter,
                    counter,
                    pulse_reg=pc["register"],
                    pulse_dtype=pc.get("dtype", "uint16"),
                    word_order=pc.get("word_order", "big"),
                    pulse_scale=pc.get("scale", 1),
                    freq=tp.get("freq", 50),
                    lagging=tp.get("pf_lagging", True),
                    settle_s=tp.get("settle_s", 3),
                    n_periods=tp.get("n_periods", 10),
                    on_point_start=_make_point_start_cb(sheet, total),
                    on_progress=_make_progress_cb(sheet),
                    should_stop=lambda: self._stop,
                )

                # Write back successful rows
                ok_results = [r for r in results if r["error"] is None]
                try:
                    if ok_results:
                        write_results(p["xlsx"], ok_results, sheet=sheet)
                except Exception as exc:
                    self.root.after(
                        0, messagebox.showwarning,
                        "写回 xlsx", f"【{sheet}】写回失败：{exc}",
                    )

            # Final status
            stopped = self._stop
            final_msg = "已停止" if stopped else "全部完成"
            self.root.after(0, self._worker_done, final_msg)

        except Exception as exc:
            self.root.after(0, self._worker_error, "测试异常", str(exc))
        finally:
            # Close all clients regardless of success/failure
            for client in (source, counter, meter):
                if client is not None:
                    try:
                        client.close()
                    except Exception:
                        pass

    # ------------------------------------------------------------------ #
    #  Main-thread helpers (all called via root.after)
    # ------------------------------------------------------------------ #

    def _ask_continue(self, sheet: str, wval: int):
        """Show wiring-confirm dialog on main thread; unblock worker when done."""
        ok = messagebox.askokcancel(
            "确认接线",
            f"电表接线方式已设为【{sheet}】(寄存器值={wval})。\n"
            "请确认/改好实际接线后点「确定」继续；点「取消」停止测试。",
        )
        self._pause_result["ok"] = ok
        self._resume.set()

    def _insert_skip_row(self, sheet: str, reason: str):
        """Insert a skip/error row into the Treeview — main thread."""
        self.tree.insert("", "end", values=(
            sheet, "", "", "", "", "", "", "", reason,
        ))
        children = self.tree.get_children()
        if children:
            self.tree.see(children[-1])

    def _on_progress(self, sheet: str, i: int, result: dict):
        """Insert one result row into the Treeview — main thread."""
        status = "OK" if result["error"] is None else result["error"]

        def _fmt(v):
            return "" if v is None else f"{v:.6g}"

        self.tree.insert("", "end", values=(
            sheet,
            result["row"],
            _fmt(result["voltage"]),
            _fmt(result["current"]),
            _fmt(result["power_factor"]),
            _fmt(result["min_s"]),
            _fmt(result["max_s"]),
            _fmt(result["avg_s"]),
            status,
        ))
        children = self.tree.get_children()
        if children:
            self.tree.see(children[-1])

    def _worker_done(self, msg: str):
        """Re-enable buttons and show completion message — main thread."""
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.lbl_prog.config(text=msg)
        messagebox.showinfo("测试完成", msg)

    def _worker_error(self, title: str, msg: str):
        """Show error dialog and re-enable buttons — main thread."""
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.lbl_prog.config(text="出错")
        messagebox.showerror(title, msg)


def main():
    root = tk.Tk()
    app = AutoTestApp(root)  # noqa: F841
    root.minsize(800, 580)
    root.mainloop()
