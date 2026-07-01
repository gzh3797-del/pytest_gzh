"""按关键词在控件模型(pages.json)里搜上位机控件 —— 给"手工用例 → 自动化"定位
页面/控件用。同时把控件的 data 名反查出对应寄存器地址(供 write_verify 定写靶、
read_compare 跨源比对)。

用法:
    python -m comm.ctl_acuview.find_widget backlight
    python -m comm.ctl_acuview.find_widget password
    python -m comm.ctl_acuview.find_widget --page General        # 列某页全部控件

匹配 page / widget 名 / data 名(大小写不敏感子串)。
"""
from __future__ import annotations

import sys

from .config import consume_config_arg
from .spec_loader import load_spec, reg_by_name

# 内容区可视高度近似(1080 - content_origin_y≈144); y 超过则该控件需滚动才能看到
_VISIBLE_H = 900


def _iter_widgets(pages):
    for page_name, page in pages["pages"].items():
        for wname, w in page.get("widgets", {}).items():
            yield page_name, wname, w


def search(keyword: str | None, page_filter: str | None) -> list[tuple]:
    registers, pages = load_spec()
    kw = (keyword or "").lower()
    out = []
    for page_name, wname, w in _iter_widgets(pages):
        if page_filter and page_filter.lower() != page_name.lower():
            continue
        data = w.get("data") or ""
        hay = f"{page_name} {wname} {data}".lower()
        if kw and kw not in hay:
            continue
        reg = reg_by_name(registers, data) if data else None
        out.append((page_name, wname, w, reg))
    out.sort(key=lambda t: (t[0], t[2].get("y") or 0))
    return out


def _fmt(page_name, wname, w, reg) -> str:
    y = w.get("y")
    scroll = " [需滚动]" if isinstance(y, int) and y > _VISIBLE_H else ""
    regtxt = f"addr={reg['addr']} ({reg['addr_hex']}) {reg['dtype']} {reg['rw']}" if reg else "(无对应寄存器/data)"
    return (f"  [{page_name}] {wname}{scroll}\n"
            f"      type={w.get('type')}  data={w.get('data')}  xy=({w.get('x')},{y}) wh=({w.get('w')},{w.get('h')})\n"
            f"      寄存器: {regtxt}")


def main(argv):
    argv = consume_config_arg(argv)
    page_filter = None
    if "--page" in argv:
        i = argv.index("--page")
        page_filter = argv[i + 1] if i + 1 < len(argv) else None
        argv = argv[:i] + argv[i + 2:]
    if not argv and not page_filter:
        print('用法: python -m comm.ctl_acuview.find_widget "<关键词>"  或  --page <页面名>')
        return 2
    kw = " ".join(argv) if argv else None
    hits = search(kw, page_filter)
    print(f'控件搜索 kw={kw!r} page={page_filter!r}: 命中 {len(hits)} 条')
    for page_name, wname, w, reg in hits[:80]:
        print(_fmt(page_name, wname, w, reg))
    if len(hits) > 80:
        print(f"  ... 还有 {len(hits) - 80} 条, 请缩小关键词或用 --page")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
