"""
接线检查测试报告生成器（HTML 格式）

表格列：电表型号 | 电表IP | 接线方式 | 用例编号 | 输入信号
         | 电压-A 预期/实测 | 电压-B 预期/实测 | 电压-C 预期/实测
         | [User1-A 预期/实测 | User1-B 预期/实测 | User1-C 预期/实测] × N通道
         | 通过与否

User Channel 数量由 actual_i 列表长度动态决定（后续按接线方式限定）。
"""
import html as _html
from datetime import datetime
from pathlib import Path

REPORT_DIR = Path(__file__).parent.parent / 'reports'
_METER_MODEL = 'AcuRev-4100'


# ── 工具 ─────────────────────────────────────────────────────────────────────

def _e(v):
    return _html.escape(str(v)) if v is not None else ''


def _fmt_src(src: dict) -> str:
    if not src:
        return ''
    lines = [
        f"Va={src.get('ua',0):.0f}V@{src.get('qua',0):.0f}°  "
        f"Vb={src.get('ub',0):.0f}V@{src.get('qub',0):.0f}°  "
        f"Vc={src.get('uc',0):.0f}V@{src.get('quc',0):.0f}°",
        f"Ia={src.get('ia',0):.2f}A@{src.get('qia',0):.0f}°  "
        f"Ib={src.get('ib',0):.2f}A@{src.get('qib',0):.0f}°  "
        f"Ic={src.get('ic',0):.2f}A@{src.get('qic',0):.0f}°",
    ]
    return '<br>'.join(_e(l) for l in lines)


def _match(exp: str, act: str) -> bool:
    e, a = exp.lower().strip(), act.lower().strip()
    if e == a:
        return True
    if e in ('pass',) and a in ('pass', 'ok', '—', '-', ''):
        return True
    for kw in ('missing', 'reversed', 'phase shift', 'phase order'):
        if kw in e and kw in a:
            return True
    # Phase Order Error：UI 可能只显示 "Error"，只要预期含 "phase order" 且实测含 "error" 即匹配
    if 'phase order' in e and 'error' in a:
        return True
    return False


def _status_cell(val: str, exp: str = None) -> str:
    v = (val or '').strip()
    if not v or v == 'N/A':
        # 实测为空：若预期值不允许空值匹配，标红提示数据未采集
        if exp is not None and not _match(exp, ''):
            return '<td style="background:#ff4444;color:#ffffff;text-align:center">—</td>'
        return '<td style="color:#aaa;text-align:center">—</td>'
    if exp is not None:
        # 有预期值：匹配→绿，不匹配→深红
        if _match(exp, v):
            bg, color = '#d4edda', '#155724'
        else:
            bg, color = '#ff4444', '#ffffff'
    else:
        # 无预期值：按 fault 类型着色
        if v.lower() == 'pass':
            bg, color = '#d4edda', '#155724'
        elif 'missing' in v.lower():
            bg, color = '#fff3cd', '#856404'
        elif 'reversed' in v.lower():
            bg, color = '#f8d7da', '#721c24'
        elif 'shift' in v.lower() or 'error' in v.lower():
            bg, color = '#fce8d5', '#7d3c00'
        else:
            bg, color = '#ffffff', '#333333'
    return (f'<td style="background:{bg};color:{color};'
            f'text-align:center;font-size:11px">{_e(v)}</td>')


# ── 样式 / JS ─────────────────────────────────────────────────────────────────

_CSS = """
body{font-family:'Microsoft YaHei',Arial,sans-serif;font-size:12px;
     margin:20px;background:#f5f5f5;color:#333}
h1{font-size:18px;margin-bottom:4px}
.meta{color:#666;font-size:11px;margin-bottom:12px}
.cards{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap}
.card{background:#fff;border-radius:6px;padding:12px 18px;
      box-shadow:0 1px 4px rgba(0,0,0,.1);text-align:center;min-width:90px}
.card .num{font-size:26px;font-weight:bold}
.card .lbl{font-size:10px;color:#888}
.card.total .num{color:#007bff}.card.pass .num{color:#28a745}.card.fail .num{color:#dc3545}
table{border-collapse:collapse;width:100%;background:#fff;
      border-radius:6px;margin-bottom:0}
.tbl-wrap{overflow-x:auto;width:100%;border-radius:6px;
          box-shadow:0 1px 4px rgba(0,0,0,.1);margin-bottom:0}
.sticky-bar{position:fixed;bottom:0;left:0;right:0;overflow-x:auto;height:14px;
            background:#e9ecef;border-top:1px solid #ced4da;z-index:100}
.tbl-wrap{margin-bottom:30px}
#sb-inner{height:1px}
thead tr{background:#343a40;color:#fff}
th{padding:7px 6px;text-align:center;font-size:11px;white-space:nowrap;
   border:1px solid #4a5060}
td{padding:5px 6px;border:1px solid #eee;vertical-align:middle;font-size:11px}
td.fixed{text-align:center;white-space:nowrap}
tr:nth-child(even){background:#fafafa}
tr:hover td{filter:brightness(.96)}
.grp-v{background:#1a5276}.grp-i{background:#1a3a5c}
.sub-exp{background:#2e6da4}.sub-act{background:#1a4a7a}
.pass-badge{background:#d4edda;color:#155724;padding:2px 8px;
            border-radius:10px;font-weight:bold;font-size:11px}
.fail-badge{background:#f8d7da;color:#721c24;padding:2px 8px;
            border-radius:10px;font-weight:bold;font-size:11px}
"""

_TOGGLE = """
<button id="btn" onclick="toggle()"
  style="position:fixed;top:16px;right:20px;z-index:999;padding:5px 14px;
         background:#dc3545;color:#fff;border:none;border-radius:4px;
         cursor:pointer;font-size:12px">仅显示异常</button>
<script>
function toggle(){
  var on=document.getElementById('btn').dataset.on==='1';
  document.querySelectorAll('tbody tr').forEach(function(tr){
    tr.style.display=(on||tr.dataset.fail!=='1')?'':'none';
  });
  var b=document.getElementById('btn');
  b.textContent=on?'仅显示异常':'显示全部';
  b.style.background=on?'#dc3545':'#6c757d';
  b.dataset.on=on?'0':'1';
}
document.addEventListener('DOMContentLoaded',function(){
  var w=document.getElementById('tw'),
      sb=document.getElementById('sb'),
      si=document.getElementById('sb-inner');
  if(!w||!sb||!si)return;
  function sync(){si.style.width=w.scrollWidth+'px';}
  sync();
  w.addEventListener('scroll',function(){sb.scrollLeft=w.scrollLeft;});
  sb.addEventListener('scroll',function(){w.scrollLeft=sb.scrollLeft;});
  window.addEventListener('resize',sync);
});
</script>
"""


# ── 主函数 ────────────────────────────────────────────────────────────────────

def generate(results: list[dict], wiring_type: str, device_name: str,
             meter_ip: str, active_channels: int = None,
             channel_phases: tuple = None,
             channel_phase_map: tuple = None,
             voltage_phases: tuple = None, **_) -> str:
    """
    channel_phases: 电流列显示的相列表，默认 ('A','B','C')。
                    1E2W 传 ('A',)，Delta/1Phase 传 ('A','C')。
    """
    REPORT_DIR.mkdir(exist_ok=True)
    ts  = datetime.now().strftime('%Y%m%d_%H%M%S')
    out = REPORT_DIR / f'wiring_{wiring_type.replace(" ", "_")}_{ts}.html'

    total  = len(results)
    passed = sum(1 for r in results if r['pass_'])
    failed = total - passed

    # 活跃 User Channel 数
    if active_channels is not None:
        max_ch = active_channels
    else:
        max_ch = max((len(r.get('actual_i') or []) for r in results), default=0)

    # 电流列相位
    ch_phases = channel_phases if channel_phases else ('A', 'B', 'C')
    # 电压列相位：优先使用显式传入的 voltage_phases，否则按电流列推断
    if voltage_phases:
        v_phases = voltage_phases
    elif ch_phases == ('A',):
        v_phases = ('A',)   # 1E2W
    else:
        v_phases = ('A', 'B', 'C', 'order')

    # ── 表头（两行）─────────────────────────────────────────────────────────
    hdr1 = (
        '<tr>'
        '<th rowspan="2">电表型号</th>'
        '<th rowspan="2">电表IP</th>'
        '<th rowspan="2">接线方式</th>'
        '<th rowspan="2">用例编号</th>'
        '<th rowspan="2">脚本ID</th>'
        '<th rowspan="2">规格参考</th>'
        '<th rowspan="2">输入信号</th>'
    )
    for ph in v_phases:
        label = 'Phase Order' if ph == 'order' else f'电压 {ph}'
        hdr1 += f'<th colspan="2" class="grp-v">{label}</th>'
    for ch in range(1, max_ch + 1):
        for ph in ch_phases:
            hdr1 += f'<th colspan="2" class="grp-i">U{ch}-{ph}</th>'
    hdr1 += '<th rowspan="2">通过</th><th rowspan="2">耗时(s)</th></tr>'

    hdr2 = '<tr>'
    for _ in range(len(v_phases) + max_ch * len(ch_phases)):
        hdr2 += '<th class="sub-exp">预期</th><th class="sub-act">实测</th>'
    hdr2 += '</tr>'

    # ── 数据行 ──────────────────────────────────────────────────────────────
    rows = []
    for r in results:
        exp_v      = r.get('expected', {}).get('voltage', {}) or {}
        exp_i      = r.get('expected', {}).get('current', {}) or {}
        act_v      = r.get('actual_v', {}) or {}
        act_i_list = r.get('actual_i', []) or []
        src        = r.get('src', {})
        badge = ('<span class="pass-badge">PASS</span>' if r['pass_']
                 else '<span class="fail-badge">FAIL</span>')
        fail_attr = ' data-fail="1"' if not r['pass_'] else ''

        spec_ref = r.get('spec_ref', '')
        tc_id    = r.get('tc_id', '')
        row = (
            f'<tr{fail_attr}>'
            f'<td class="fixed">{_e(_METER_MODEL)}</td>'
            f'<td class="fixed">{_e(meter_ip)}</td>'
            f'<td class="fixed">{_e(wiring_type)}</td>'
            f'<td class="fixed" style="font-size:10px;color:#1a3a5c">{_e(tc_id)}</td>'
            f'<td class="fixed">{_e(r["id"])}</td>'
            f'<td class="fixed" style="color:#555;font-size:10px">{_e(spec_ref)}</td>'
            f'<td style="white-space:nowrap">{_fmt_src(src)}</td>'
        )
        for ph in v_phases:
            ev = exp_v.get(ph, '')
            av = act_v.get(ph, '')
            row += f'<td style="text-align:center;font-size:11px">{_e(ev)}</td>'
            row += _status_cell(av, ev)
        for ch_idx in range(max_ch):
            act_i = act_i_list[ch_idx] if ch_idx < len(act_i_list) else {}
            assigned = (channel_phase_map[ch_idx]
                        if channel_phase_map and ch_idx < len(channel_phase_map)
                        else None)
            for ph in ch_phases:
                if assigned is not None and ph not in assigned:
                    # 该 channel 未分配此相，预期和实测均显示为灰色 N/A
                    row += '<td style="color:#aaa;text-align:center;font-size:11px">N/A</td>'
                    row += '<td style="color:#aaa;text-align:center">—</td>'
                    continue
                ei = exp_i.get(ph, '')
                ai = (act_i or {}).get(ph, '')
                row += f'<td style="text-align:center;font-size:11px">{_e(ei)}</td>'
                row += _status_cell(ai, ei)
        elapsed = r.get('elapsed', '')
        row += (f'<td class="fixed">{badge}</td>'
                f'<td class="fixed" style="color:#666">{_e(elapsed)}</td></tr>')
        rows.append(row)

    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="UTF-8">
<title>接线检查报告 — {_e(wiring_type)}</title>
<style>{_CSS}</style></head>
<body>
{_TOGGLE}
<h1>接线检查测试报告</h1>
<div class="meta">
  接线方式：{_e(wiring_type)} &nbsp;|&nbsp;
  设备：{_e(device_name)} ({_e(_METER_MODEL)}) &nbsp;|&nbsp;
  Meter IP：{_e(meter_ip)} &nbsp;|&nbsp;
  生成时间：{now_str}
</div>
<div class="cards">
  <div class="card total"><div class="num">{total}</div><div class="lbl">总计</div></div>
  <div class="card pass"><div class="num">{passed}</div><div class="lbl">PASS</div></div>
  <div class="card fail"><div class="num">{failed}</div><div class="lbl">FAIL</div></div>
</div>
<div class="tbl-wrap" id="tw"><table>
<thead>{hdr1}{hdr2}</thead>
<tbody>{''.join(rows)}</tbody>
</table></div>
<div class="sticky-bar" id="sb"><div id="sb-inner"></div></div>
</body></html>"""

    out.write_text(html_content, encoding='utf-8')
    return str(out)
