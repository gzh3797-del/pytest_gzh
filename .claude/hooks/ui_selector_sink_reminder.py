#!/usr/bin/env python3
"""PostToolUse hook：UI 探查子代理回报后，提醒主 AI 沉淀选择器。

挂在 PostToolUse 事件、matcher 匹配 Task/Agent 工具——即子代理结束、
其回报刚进入主对话的时刻触发，提醒注入到主 AI 收到的工具结果旁边。
（不要用 SubagentStop：该事件的 additionalContext 注入的是子代理自己的
对话，主 AI 看不到，还会打断子代理的收尾回复。）

判据（命中任一即提醒，其余子代理静默不产生噪音）：
1. tool_input.subagent_type 是UI 测试工程师（ui-test-engineer）
2. 子代理回报文本中出现 Playwright MCP 浏览器工具标记
   （mcp__plugin_playwright_playwright__browser_*）

沉淀目标（见 CLAUDE.md 约定 #11 / conventions.md「选择器沉淀复用」）：
    <项目知识库根>/requirements/context/<UI一级模块目录名>_context.md
"""
import json
import sys

PLAYWRIGHT_MARKER = "mcp__plugin_playwright_playwright__browser_"
UI_AGENT_TYPES = ("ui-test-engineer",)

REMINDER = (
    "【选择器沉淀提醒】刚回报的子 agent 做了 UI 页面探查（UI 测试工程师或用过 "
    "Playwright 浏览器工具）。若它回报了新探明的选择器/交互，请按 CLAUDE.md 约定 #11 "
    "将其沉淀到该项目的 `<项目知识库根>/requirements/context/<UI一级模块目录名>_context.md`"
    "（命中已有文档则增量合并、去重），并在项目 context.md 就近补一行指针。"
    "若本次无新选择器或已覆盖，忽略本提醒。"
)


def _hit(data: dict) -> bool:
    tool_input = data.get("tool_input")
    if isinstance(tool_input, dict):
        subagent_type = str(tool_input.get("subagent_type", "")).strip()
        if subagent_type in UI_AGENT_TYPES:
            return True
    response_blob = json.dumps(data.get("tool_response"), ensure_ascii=False, default=str)
    return PLAYWRIGHT_MARKER in response_blob


def main() -> None:
    try:
        data = json.load(sys.stdin)
    except (ValueError, OSError):
        return
    if not isinstance(data, dict) or not _hit(data):
        return  # 非 UI 探查 → 静默
    out = {
        "hookSpecificOutput": {
            "hookEventName": "PostToolUse",
            "additionalContext": REMINDER,
        },
        "suppressOutput": True,
    }
    # 用 ensure_ascii=True（默认）转义中文为 \uXXXX，规避 Windows 管道 GBK 编码问题；
    # 解析端（Claude Code）会把转义还原成正确中文。
    print(json.dumps(out))


if __name__ == "__main__":
    main()
