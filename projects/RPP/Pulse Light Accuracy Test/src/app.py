# src/app.py
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from src.sdg_client import SDGClient
from src.transport import SocketTransport, VisaTransport, list_usb_resources
from src.counter import acquire, NoPulseError
from src.storage import row_from_result, append_csv, row_to_tsv, STAT_FIELDS

_FIELD_CN = {"period": "周期(s)", "pw": "正脉宽(s)", "duty": "占空比(%)",
             "frqdev": "频偏(ppm)", "frq": "频率(Hz)"}
_ROWS = [("value", "Value"), ("mean", "Mean"), ("min", "Min"),
         ("max", "Max"), ("sdev", "Sdev")]

class App:
    def __init__(self, root):
        self.root = root
        self.client = None
        self.polling = False
        self._poll_gen = 0
        self.last_config = {}
        root.title("SDG 2042X 频率计取数")
        self._build_conn(root)
        self._build_config(root)
        self._build_data(root)

    # ---- 连接区 ----
    def _build_conn(self, root):
        f = ttk.LabelFrame(root, text="连接")
        f.pack(fill="x", padx=8, pady=4)
        top = ttk.Frame(f); top.pack(fill="x")
        ttk.Label(top, text="方式:").pack(side="left", padx=4)
        self.v_conn = tk.StringVar(value="网口")
        cb = ttk.Combobox(top, textvariable=self.v_conn, values=["网口", "USB"],
                          width=6, state="readonly")
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self._toggle_conn())
        self.btn_conn = ttk.Button(top, text="连接", command=self.on_connect)
        self.btn_conn.pack(side="left", padx=6)
        self.lbl_status = ttk.Label(top, text="● 未连接", foreground="red")
        self.lbl_status.pack(side="left")

        # 网口输入
        self.frm_lan = ttk.Frame(f)
        ttk.Label(self.frm_lan, text="IP:").pack(side="left", padx=4)
        self.ip = ttk.Entry(self.frm_lan, width=16); self.ip.pack(side="left")
        self.ip.insert(0, "192.168.1.100")
        ttk.Label(self.frm_lan, text="端口:").pack(side="left", padx=4)
        self.port = ttk.Entry(self.frm_lan, width=6); self.port.pack(side="left")
        self.port.insert(0, "5024")

        # USB 输入
        self.frm_usb = ttk.Frame(f)
        ttk.Label(self.frm_usb, text="USB资源:").pack(side="left", padx=4)
        self.v_usb = tk.StringVar(value="")
        self.cb_usb = ttk.Combobox(self.frm_usb, textvariable=self.v_usb, width=40)
        self.cb_usb.pack(side="left")
        self.btn_scan = ttk.Button(self.frm_usb, text="扫描", command=self.on_scan_usb)
        self.btn_scan.pack(side="left", padx=4)

        self._toggle_conn()

    def _toggle_conn(self):
        if self.v_conn.get() == "USB":
            self.frm_lan.pack_forget()
            self.frm_usb.pack(fill="x", pady=2)
        else:
            self.frm_usb.pack_forget()
            self.frm_lan.pack(fill="x", pady=2)

    # ---- 配置区 ----
    def _build_config(self, root):
        f = ttk.LabelFrame(root, text="频率计配置")
        f.pack(fill="x", padx=8, pady=4)
        self.v_counter = tk.StringVar(value="ON")
        self.v_mode = tk.StringVar(value="DC")
        self.v_hfr = tk.StringVar(value="ON")
        ttk.Label(f, text="计数器:").grid(row=0, column=0, sticky="e")
        ttk.Combobox(f, textvariable=self.v_counter, values=["ON", "OFF"],
                     width=5, state="readonly").grid(row=0, column=1)
        ttk.Label(f, text="模式:").grid(row=0, column=2, sticky="e")
        ttk.Combobox(f, textvariable=self.v_mode, values=["AC", "DC"],
                     width=5, state="readonly").grid(row=0, column=3)
        ttk.Label(f, text="高频抑制:").grid(row=0, column=4, sticky="e")
        ttk.Combobox(f, textvariable=self.v_hfr, values=["ON", "OFF"],
                     width=5, state="readonly").grid(row=0, column=5)
        ttk.Label(f, text="参考频率Hz:").grid(row=1, column=0, sticky="e")
        self.e_refq = ttk.Entry(f, width=10); self.e_refq.grid(row=1, column=1)
        self.e_refq.insert(0, "1000")
        ttk.Label(f, text="触发电平V:").grid(row=1, column=2, sticky="e")
        self.e_trg = ttk.Entry(f, width=10); self.e_trg.grid(row=1, column=3)
        self.e_trg.insert(0, "1.5")
        # Fix 3: keep reference to btn_apply for disable/re-enable
        self.btn_apply = ttk.Button(f, text="应用配置", command=self.on_apply)
        self.btn_apply.grid(row=1, column=5, padx=4)

    # ---- 数据区 ----
    def _build_data(self, root):
        f = ttk.LabelFrame(root, text="数据")
        f.pack(fill="both", expand=True, padx=8, pady=4)
        cols = ["stat"] + STAT_FIELDS
        self.tree = ttk.Treeview(f, columns=cols, show="headings", height=6)
        self.tree.heading("stat", text="")
        self.tree.column("stat", width=60, anchor="center")
        for fld in STAT_FIELDS:
            self.tree.heading(fld, text=_FIELD_CN[fld])
            self.tree.column(fld, width=110, anchor="center")
        self.tree.pack(fill="x")
        self.row_ids = {}
        for key, label in _ROWS:
            self.row_ids[key] = self.tree.insert("", "end",
                values=[label] + [""] * len(STAT_FIELDS))
        self.row_ids["num"] = self.tree.insert("", "end",
            values=["Num"] + [""] * len(STAT_FIELDS))

        bar = ttk.Frame(f); bar.pack(fill="x", pady=4)
        ttk.Label(bar, text="采集周期数N:").pack(side="left")
        self.e_n = ttk.Entry(bar, width=6); self.e_n.pack(side="left")
        self.e_n.insert(0, "10")
        self.btn_grab = ttk.Button(bar, text="取数", command=self.on_grab)
        self.btn_grab.pack(side="left", padx=6)
        ttk.Button(bar, text="清除统计", command=self.on_clear).pack(side="left")
        self.lbl_prog = ttk.Label(bar, text=""); self.lbl_prog.pack(side="left", padx=8)

        pathbar = ttk.Frame(f); pathbar.pack(fill="x")
        ttk.Label(pathbar, text="CSV:").pack(side="left")
        self.e_csv = ttk.Entry(pathbar, width=48); self.e_csv.pack(side="left", fill="x", expand=True)
        self.e_csv.insert(0, "counter_log.csv")
        ttk.Button(pathbar, text="...", width=3, command=self.on_pick_csv).pack(side="left")

    # ---- 行为 ----

    # Fix 1: move SDGClient construction + connect() off the main thread
    def on_scan_usb(self):
        self.btn_scan.config(state="disabled")
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self):
        try:
            res = list_usb_resources()
            self.root.after(0, self._scan_done, res, None)
        except Exception as e:
            self.root.after(0, self._scan_done, None, str(e))

    def _scan_done(self, res, err):
        self.btn_scan.config(state="normal")
        if err:
            messagebox.showerror("扫描失败", "需要已安装 NI-VISA：\n" + err)
            return
        self.cb_usb["values"] = res
        if res:
            self.cb_usb.current(0)
        else:
            messagebox.showinfo("扫描", "未发现 USB 设备")

    def on_connect(self):
        try:
            if self.v_conn.get() == "USB":
                resource = self.v_usb.get().strip()
                if not resource:
                    messagebox.showwarning("连接", "请先扫描并选择 USB 资源")
                    return
                transport = VisaTransport(resource)
            else:
                transport = SocketTransport(self.ip.get().strip(), int(self.port.get()))
        except ValueError:
            messagebox.showerror("连接", "端口必须是整数")
            return
        self.btn_conn.config(state="disabled")
        threading.Thread(target=self._connect_worker, args=(transport,), daemon=True).start()

    def _connect_worker(self, transport):
        try:
            client = SDGClient(transport)
            client.connect()
            self.root.after(0, self._on_connect_ok, client)
        except Exception as e:
            self.root.after(0, self._on_connect_fail, str(e))

    def _on_connect_ok(self, client):
        self.client = client
        self.lbl_status.config(text="● 已连接", foreground="green")
        self.btn_conn.config(state="normal")
        self.start_polling()

    def _on_connect_fail(self, msg):
        self.btn_conn.config(state="normal")
        messagebox.showerror("连接失败", msg)

    # Fix 3: move set_* calls off the main thread
    def on_apply(self):
        if not self._require_client():
            return
        # Snapshot client and GUI values on main thread before handing off to worker
        client = self.client
        counter_on = self.v_counter.get() == "ON"
        mode = self.v_mode.get()
        hfr_on = self.v_hfr.get() == "ON"
        refq = self.e_refq.get().strip()
        trg = self.e_trg.get().strip()
        config_snapshot = {"mode": mode, "hfr": self.v_hfr.get(),
                           "refq": refq, "trg": trg}
        self.btn_apply.config(state="disabled")
        threading.Thread(
            target=self._apply_worker,
            args=(client, counter_on, mode, hfr_on, refq, trg, config_snapshot),
            daemon=True,
        ).start()

    def _apply_worker(self, client, counter_on, mode, hfr_on, refq, trg, config_snapshot):
        try:
            client.set_counter(counter_on)
            client.set_mode(mode)
            client.set_hfr(hfr_on)
            client.set_refq(refq)
            client.set_trg(trg)
        except Exception as e:
            self.root.after(0, self._on_apply_done, None, str(e))
            return
        self.root.after(0, self._on_apply_done, config_snapshot, None)

    def _on_apply_done(self, config_snapshot, err):
        self.btn_apply.config(state="normal")
        if err:
            messagebox.showerror("配置失败", err)
        else:
            self.last_config = config_snapshot
            messagebox.showinfo("配置", "已下发配置")

    # Fix 2: generation token ensures only the newest poll loop survives reconnect
    def start_polling(self):
        self._poll_gen += 1
        gen = self._poll_gen
        self.polling = True
        threading.Thread(target=self._poll_loop, args=(gen,), daemon=True).start()

    def _poll_loop(self, gen):
        consecutive_errors = 0
        while self.polling and self.client and gen == self._poll_gen:
            try:
                d = self.client.query_fcnt()
                consecutive_errors = 0
                self.root.after(0, self._update_value_row, d)
            except Exception:
                consecutive_errors += 1
                if consecutive_errors >= 5:
                    self.root.after(0, self._on_poll_lost)
                    return
            time.sleep(0.7)

    # Fix 2: called on main thread when poll loses connection
    def _on_poll_lost(self):
        self.polling = False
        self.client = None
        self.lbl_status.config(text="● 连接丢失", foreground="red")

    def _update_value_row(self, d):
        vals = ["Value"]
        for fld in STAT_FIELDS:
            v = d.get(fld)
            vals.append("" if v is None else f"{v:.6g}")
        self.tree.item(self.row_ids["value"], values=vals)

    def on_grab(self):
        # Fix 5: re-entrancy guard
        if self.btn_grab["state"] == "disabled":
            return
        if not self._require_client():
            return
        try:
            n = int(self.e_n.get())
        except ValueError:
            messagebox.showerror("取数", "周期数N必须是整数")
            return
        client = self.client
        self.btn_grab.config(state="disabled")
        self.lbl_prog.config(text="采集中…")
        threading.Thread(target=self._grab_worker, args=(client, n), daemon=True).start()

    def _grab_worker(self, client, n):
        try:
            result = acquire(client, n_periods=n)
        except NoPulseError as e:
            self.root.after(0, self._grab_done, None, str(e))
            return
        except Exception as e:
            self.root.after(0, self._grab_done, None, "通信错误: " + str(e))
            return
        self.root.after(0, self._grab_done, result, None)

    def _grab_done(self, result, err):
        self.btn_grab.config(state="normal")
        self.lbl_prog.config(text="")
        if err:
            messagebox.showwarning("取数", err)
            return
        stats = result["stats"]
        # Fix 4: use label directly from _ROWS iteration, no dict(_ROWS) lookup
        for key, label in _ROWS:
            vals = [label]
            for fld in STAT_FIELDS:
                v = stats[fld][key]
                vals.append("" if v is None else f"{v:.6g}")
            self.tree.item(self.row_ids[key], values=vals)
        self.tree.item(self.row_ids["num"],
                       values=["Num"] + [stats["num"]] * len(STAT_FIELDS))
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        row = row_from_result(ts, result, self.last_config)
        try:
            append_csv(self.e_csv.get().strip(), row)
        except Exception as e:
            messagebox.showwarning("CSV", "写 CSV 失败: " + str(e))
        self.root.clipboard_clear()
        self.root.clipboard_append(row_to_tsv(row))
        self.lbl_prog.config(text=f"已取数 {stats['num']} 个样本，已存CSV+剪贴板")

    def on_clear(self):
        for key, label in _ROWS:
            self.tree.item(self.row_ids[key], values=[label] + [""] * len(STAT_FIELDS))
        self.tree.item(self.row_ids["num"], values=["Num"] + [""] * len(STAT_FIELDS))

    def on_pick_csv(self):
        p = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV", "*.csv")])
        if p:
            self.e_csv.delete(0, "end")
            self.e_csv.insert(0, p)

    def _require_client(self):
        if not self.client:
            messagebox.showwarning("未连接", "请先连接仪器")
            return False
        return True

def main():
    root = tk.Tk()
    App(root)
    root.minsize(720, 360)
    root.mainloop()
