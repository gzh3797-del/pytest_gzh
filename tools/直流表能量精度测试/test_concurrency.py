# -*- coding: utf-8 -*-
"""验证：并发调用时命令不再错配，且过期残包被丢弃。"""
import socket
import threading
import time

from xl9600 import XL9600, build_command, ENCODING

HOST = "127.0.0.1"


def start_mock(delay=0.05):
    """模拟设备：每收到 <X>... 就延迟后回 <X应答>本命令名;<End>"""
    srv = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    srv.bind((HOST, 0))
    port = srv.getsockname()[1]

    def loop():
        srv.settimeout(3)
        while True:
            try:
                data, addr = srv.recvfrom(8192)
            except socket.timeout:
                return
            text = data.decode(ENCODING, "replace")
            name = text[text.find("<") + 1:text.find(">")]
            time.sleep(delay)  # 放大竞态窗口
            reply = f"<{name}应答>\r\n回显:{name};\r\n<End>\r\n".encode(ENCODING)
            srv.sendto(reply, addr)

    threading.Thread(target=loop, daemon=True).start()
    return port


def test_no_mismatch():
    port = start_mock()
    dev = XL9600(HOST, port, timeout=3).open()
    results = {}
    names = ["源输出", "误差读取", "源停止", "参数配置", "日计时误差读取"]

    def call(nm):
        r = dev.send_raw(nm)
        results[nm] = r.get("回显")

    threads = [threading.Thread(target=call, args=(n,)) for n in names]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    dev.close()

    for n in names:
        assert results[n] == n, f"命令错配: {n} 收到了 {results[n]}"
    print("[PASS] 并发 5 条命令，每条都拿到自己的正确回复:", results)


def test_drain_stale():
    port = start_mock(delay=0)
    dev = XL9600(HOST, port, timeout=2).open()
    # 手动塞一个过期回复进 socket 缓冲区
    dev._sock.sendto(build_command("陈旧"), (HOST, port))
    time.sleep(0.2)  # 让 mock 把过期回复发回来，堆在缓冲区
    r = dev.send_raw("源停止")
    dev.close()
    assert r.get("_header") == "源停止应答", f"读到了过期残包: {r}"
    print("[PASS] 过期残包被丢弃，源停止 收到正确应答:", r.get("_header"))


if __name__ == "__main__":
    test_no_mismatch()
    test_drain_stale()
    print("全部通过 ✓")
