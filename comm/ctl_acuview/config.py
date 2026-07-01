"""配置加载。

读取项目根目录的 config.yaml，提供点号访问的配置对象。
所有路径相对 config.yaml 所在目录解析。
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml

# 引擎根 = 本文件上一级(comm/ctl_acuview/ 的父目录, 即 comm/)；仅作无显式 config 时的兜底
ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG_PATH = ROOT / "config.yaml"


class _Node(dict):
    """支持属性访问的 dict（cfg.transport.tcp.host）。"""

    def __getattr__(self, name: str) -> Any:
        try:
            val = self[name]
        except KeyError as exc:  # pragma: no cover - 防御性
            raise AttributeError(name) from exc
        if isinstance(val, dict) and not isinstance(val, _Node):
            val = _wrap(val)
            self[name] = val
        return val


def _wrap(obj: Any) -> Any:
    if isinstance(obj, dict):
        return _Node({k: _wrap(v) for k, v in obj.items()})
    if isinstance(obj, list):
        return [_wrap(v) for v in obj]
    return obj


class Config:
    """加载后的配置 + 路径解析辅助。"""

    def __init__(self, path: str | os.PathLike | None = None):
        self.path = Path(path) if path else DEFAULT_CONFIG_PATH
        with open(self.path, "r", encoding="utf-8") as fh:
            raw = yaml.safe_load(fh)
        self.data: _Node = _wrap(raw)
        self.root = self.path.resolve().parent

    def __getattr__(self, name: str) -> Any:
        # 代理到底层数据节点
        return getattr(self.data, name)

    def resolve(self, rel: str) -> Path:
        """把配置里的相对路径解析为绝对路径。"""
        p = Path(rel)
        return p if p.is_absolute() else (self.root / p)

    # ---- 常用派生路径 ----
    @property
    def spec_dir(self) -> Path:
        return self.resolve(self.data["spec"]["out_dir"])

    @property
    def excel_path(self) -> Path:
        return self.resolve(self.data["spec"]["excel"])

    @property
    def json_path(self) -> Path:
        return self.resolve(self.data["spec"]["json"])

    @property
    def registers_json(self) -> Path:
        return self.spec_dir / "registers.json"

    @property
    def pages_json(self) -> Path:
        return self.spec_dir / "pages.json"

    @property
    def report_dir(self) -> Path:
        return self.resolve(self.data["run"]["report_dir"])


_singleton: Config | None = None


def get_config(path: str | os.PathLike | None = None) -> Config:
    """获取(缓存的)配置单例。

    多项目: 配置按项目放在 projects/<X>/config_acuview.yaml。未显式传 path 且尚未初始化时，
    回退读环境变量 ACUVIEW_CONFIG；都没有才用 DEFAULT_CONFIG_PATH(仓库根 config.yaml)。
    """
    global _singleton
    if path is None and _singleton is None:
        env = os.environ.get("ACUVIEW_CONFIG")
        if env:
            path = env
    if _singleton is None or path is not None:
        _singleton = Config(path)
    return _singleton


def consume_config_arg(argv: list[str]) -> list[str]:
    """从命令行 argv 取出 `--config/-c <路径>` 并据此切换配置单例；返回去掉该项后的 argv。

    供各 CLI(spec_loader / find_register / find_widget) 复用，让多项目下能显式指定
    projects/<X>/config_acuview.yaml。"""
    out = list(argv)
    for flag in ("--config", "-c"):
        if flag in out:
            i = out.index(flag)
            if i + 1 < len(out):
                get_config(out[i + 1])
                del out[i:i + 2]
            else:
                del out[i]
    return out


if __name__ == "__main__":
    cfg = get_config()
    print("config path :", cfg.path)
    print("app exe     :", cfg.app.exe_path)
    print("gui transport:", cfg.transport.gui, "| verify:", cfg.transport.verify)
    print("tcp         :", dict(cfg.transport.tcp))
    print("rtu         :", dict(cfg.transport.rtu))
    print("excel       :", cfg.excel_path)
    print("json        :", cfg.json_path)
    print("spec_dir    :", cfg.spec_dir)
