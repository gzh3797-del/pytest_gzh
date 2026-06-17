from pathlib import Path

from framework.config.loader import deep_merge, load_config


def test_deep_merge_overrides_nested_keys():
    base = {"ssh": {"user": "admin", "port": 22}, "relay": {"ip": "192.168.1.100"}}
    override = {"ssh": {"user": "root"}, "device_ip": "20.20.20.20"}
    merged = deep_merge(base, override)
    assert merged["ssh"]["user"] == "root"      # override 生效
    assert merged["ssh"]["port"] == 22           # 未覆盖项保留
    assert merged["relay"]["ip"] == "192.168.1.100"
    assert merged["device_ip"] == "20.20.20.20"  # 新增项加入


def test_load_config_merges_global_then_project(tmp_path: Path):
    (tmp_path / "configs").mkdir()
    (tmp_path / "configs" / "global.yaml").write_text(
        "ssh:\n  user: admin\nrelay:\n  ip: 192.168.1.100\n", encoding="utf-8")
    project_dir = tmp_path / "projects" / "AcuHMI_1_7"
    project_dir.mkdir(parents=True)
    (project_dir / "config.yaml").write_text(
        "project_name: AcuHMI 1.7\ndevice_ip: 20.20.20.20\n", encoding="utf-8")
    cfg = load_config("AcuHMI_1_7", repo_root=tmp_path)
    assert cfg["project_name"] == "AcuHMI 1.7"
    assert cfg["ssh"]["user"] == "admin"
    assert cfg["device_ip"] == "20.20.20.20"
