"""按关键词在寄存器表里搜寄存器 —— 给"手工用例 → 自动化"定位写靶/校验地址用。

用法:
    python -m comm.ctl_acuview.find_register backlight
    python -m comm.ctl_acuview.find_register "slave id"
    python -m comm.ctl_acuview.find_register 4120          # 直接按地址(十进制)精确查

在 name(UPPER_SNAKE) 和 description(人类描述) 上做大小写不敏感子串匹配。
"""
from __future__ import annotations

import sys

from .config import consume_config_arg
from .spec_loader import load_spec, reg_by_addr


def search(keyword: str) -> list[dict]:
    registers, _ = load_spec()
    by_addr = registers["by_addr"]
    # 纯数字 -> 按地址精确
    if keyword.strip().isdigit():
        e = reg_by_addr(registers, int(keyword))
        return [e] if e else []
    kw = keyword.lower()
    hits = []
    for e in by_addr.values():
        if kw in e["name"].lower() or kw in e["description"].lower():
            hits.append(e)
    hits.sort(key=lambda e: e["addr"])
    return hits


def _fmt(e: dict) -> str:
    return (f"  addr={e['addr']} ({e['addr_hex']})  {e['dtype']:6s} {e['rw']:4s} "
            f"regN={e['reg_num']}  range={e['range'] or '-'}\n"
            f"      name={e['name']}  | sheet={e['sheet']}/{e['block']}\n"
            f"      desc={e['description']}")


def main(argv):
    argv = consume_config_arg(argv)
    if not argv:
        print('用法: python -m comm.ctl_acuview.find_register [--config <项目>/config_acuview.yaml] "<关键词或十进制地址>"')
        return 2
    kw = " ".join(argv)
    hits = search(kw)
    print(f'寄存器搜索 "{kw}": 命中 {len(hits)} 条')
    for e in hits[:50]:
        print(_fmt(e))
    if len(hits) > 50:
        print(f"  ... 还有 {len(hits) - 50} 条, 请用更具体的关键词")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
