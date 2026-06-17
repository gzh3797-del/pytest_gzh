"""分层配置加载器：configs/global.yaml <- projects/<project>/config.yaml <- .env（敏感值）。"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml

_REPO_ROOT = Path(__file__).resolve().parents[2]


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    """递归合并：override 的标量覆盖 base，dict 递归合并，base 未触及项保留。"""
    result = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def _read_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def load_config(project: str, *, repo_root: Path | None = None) -> dict[str, Any]:
    """合并全局与项目配置：configs/global.yaml ← projects/<project>/config.yaml；
    同名环境变量（大写）覆盖顶层标量。"""
    root = repo_root or _REPO_ROOT
    merged = deep_merge(
        _read_yaml(root / "configs" / "global.yaml"),
        _read_yaml(root / "projects" / project / "config.yaml"),
    )
    for key in list(merged.keys()):
        env_value = os.getenv(key.upper())
        if env_value is not None and not isinstance(merged[key], dict):
            merged[key] = env_value
    return merged
