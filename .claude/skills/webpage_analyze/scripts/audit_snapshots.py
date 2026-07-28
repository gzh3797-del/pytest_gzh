#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""穷尽性快照审计 — 跨平台替代 `ls form_* subpage_* | wc -l`。

用法:
    python audit_snapshots.py <output_dir> [expected_ops.json] [ctx_dir]

行为:
    1. 按前缀统计 <output_dir>（临时工作目录）下的探索中间快照文件数
       （form_select_/form_checkbox_/form_radio_/form_submit_/form_reset_/
       form_search_/subpage_）。
    2. 若给出 expected_ops.json（或含 ```json``` 块的报告），推导各维度预期数：
         - 优先 expected_counts（若含该字段）
         - 否则用 element_counts（form 页真实 schema）
         - 子页面预期数恒为 len(sub_pages)
    3. 输出逐维度「预期 vs 实际」表。任一维度 实际 < 预期 → 退出码 1。
       无预期数据时仅打印实际计数（信息性，退出码 0）。
    4. 若给出 ctx_dir（交付目录 requirements/context/），额外**信息性**报告已交付的
       `*_context.md` 数与 `_INDEX_context.md` 是否存在（代表性合并会使文件数少于
       子页面数，故不作硬性 FAIL 门禁）。

Windows / PowerShell / bash 通用，仅依赖标准库。
"""
import sys
import os
import re
import json
import glob

PREFIXES = {
    "下拉框":   "form_select_*.md",
    "复选框":   "form_checkbox_*.md",
    "单选组":   "form_radio_*.md",
    "提交按钮": "form_submit_*.md",
    "重置按钮": "form_reset_*.md",
    "搜索框":   "form_search_*.md",
    "子页面":   "subpage_*.md",
}


def count_actual(output_dir):
    counts = {}
    for dim, pat in PREFIXES.items():
        counts[dim] = len(glob.glob(os.path.join(output_dir, pat)))
    return counts


def load_json_block(path):
    """接受两种预期来源:
    - .json 文件（如阶段1产出的 expected_ops.json）→ 直接解析
    - .md 报告 → 提取 ```json``` 块
    """
    with open(path, encoding="utf-8") as f:
        content = f.read()
    if path.lower().endswith(".json"):
        return json.loads(content)
    m = re.search(r"```json\n(.*?)\n```", content, re.DOTALL)
    if not m:
        return None
    return json.loads(m.group(1))


def expected_from_json(data):
    """从报告 JSON 推导各维度预期快照数。容忍多种真实 schema。"""
    exp = {}
    ec = data.get("expected_counts") or {}
    el = data.get("element_counts") or {}

    # 下拉框：expected_counts.total_options 优先，否则 element_counts.select
    if "total_options" in ec:
        exp["下拉框"] = ec["total_options"]
    elif "select" in el:
        exp["下拉框"] = el["select"]

    # 复选框 ×2 状态
    if "checkboxes" in ec:
        exp["复选框"] = ec["checkboxes"] * 2
    elif "input_checkbox" in el:
        exp["复选框"] = el["input_checkbox"] * 2

    if "radios" in ec:
        exp["单选组"] = ec["radios"]
    elif "input_radio" in el:
        exp["单选组"] = el["input_radio"]

    if "submit_buttons" in ec:
        exp["提交按钮"] = ec["submit_buttons"]
    elif "input_submit" in el:
        exp["提交按钮"] = el["input_submit"]

    if "reset_buttons" in ec:
        exp["重置按钮"] = ec["reset_buttons"]
    elif "input_reset" in el:
        exp["重置按钮"] = el["input_reset"]

    subs = data.get("sub_pages")
    if isinstance(subs, list):
        exp["子页面"] = len(subs)

    return exp


def main():
    if len(sys.argv) < 2:
        print("用法: python audit_snapshots.py <output_dir> [expected_ops.json] [ctx_dir]")
        return 2

    output_dir = sys.argv[1]
    if not os.path.isdir(output_dir):
        print(f"ERROR: 产物目录不存在: {output_dir}")
        return 2

    actual = count_actual(output_dir)
    expected = {}
    if len(sys.argv) >= 3:
        report = sys.argv[2]
        if not os.path.isfile(report):
            print(f"ERROR: 报告文件不存在: {report}")
            return 2
        data = load_json_block(report)
        if data is None:
            print("WARN: 报告缺少 ```json``` 块，仅输出实际计数。")
        else:
            expected = expected_from_json(data)

    print(f"产物目录: {output_dir}")
    print(f"{'维度':<8}{'预期':>6}{'实际':>6}  状态")
    print("-" * 32)
    failed = []
    for dim in PREFIXES:
        a = actual.get(dim, 0)
        if dim in expected:
            e = expected[dim]
            ok = a >= e
            status = "PASS" if ok else "FAIL"
            if not ok:
                failed.append(f"{dim}: 预期>={e}, 实际={a}")
            print(f"{dim:<8}{e:>6}{a:>6}  {status}")
        else:
            print(f"{dim:<8}{'-':>6}{a:>6}  (无预期)")

    print("-" * 32)

    if len(sys.argv) >= 4:
        ctx_dir = sys.argv[3]
        if os.path.isdir(ctx_dir):
            ctx_files = [
                p for p in glob.glob(os.path.join(ctx_dir, "*_context.md"))
                if os.path.basename(p) != "_INDEX_context.md"
            ]
            has_index = os.path.isfile(os.path.join(ctx_dir, "_INDEX_context.md"))
            print(f"交付目录: {ctx_dir}")
            print(f"  已交付 *_context.md: {len(ctx_files)} 份"
                  "（代表性合并会少于子页面数，仅信息性）")
            print(f"  _INDEX_context.md: {'存在' if has_index else '缺失 ✗'}")
        else:
            print(f"WARN: 交付目录不存在: {ctx_dir}")
        print("-" * 32)

    if not expected:
        print("VERDICT: NO_EXPECTED (仅信息性计数)")
        return 0
    if failed:
        print("VERDICT: FAIL")
        for f in failed:
            print("  ✗ " + f)
        return 1
    print("VERDICT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
