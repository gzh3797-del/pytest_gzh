from pathlib import Path

from framework.report.paths import detect_module, make_report_dir, update_latest


def test_make_report_dir_creates_subdirs(tmp_path: Path):
    dirs = make_report_dir("AcuHMI_1_7", "20260610_153000", reports_root=tmp_path)
    assert dirs.root == tmp_path / "AcuHMI_1_7" / "20260610_153000"
    assert (dirs.root).is_dir()
    assert (dirs.screenshots).is_dir()
    # report.html 直接落根目录，不再有 html/ 子目录；logs/ 不再创建
    assert not (dirs.root / "html").exists()
    assert not (dirs.root / "logs").exists()


def test_make_report_dir_inserts_module_level(tmp_path: Path):
    dirs = make_report_dir("AcuHMI_1_7", "20260610_153000",
                           module="BacnetIP", reports_root=tmp_path)
    assert dirs.root == tmp_path / "AcuHMI_1_7" / "BacnetIP" / "20260610_153000"
    assert (dirs.root).is_dir()
    assert (dirs.screenshots).is_dir()


def test_detect_module_single_module():
    args = ["projects/AcuHMI_1_7/tests/BacnetIP/test_x.py::TestC::test_m"]
    assert detect_module(args) == "BacnetIP"


def test_detect_module_nested_returns_top_level():
    args = ["projects/AcuHMI_1_7/tests/ui/protocols/snmp/test_x.py"]
    assert detect_module(args) == "ui"


def test_detect_module_windows_separators():
    args = [r"projects\acuhmi_1_7\tests\wiring_check\test_x.py"]
    assert detect_module(args) == "Wiring_check"


def test_detect_module_ambiguous_returns_none():
    args = ["projects/AcuHMI_1_7/tests/BacnetIP", "projects/AcuHMI_1_7/tests/ui"]
    assert detect_module(args) is None


def test_detect_module_whole_project_returns_none():
    assert detect_module(["projects/AcuHMI_1_7"]) is None
    assert detect_module(["-m", "smoke"]) is None
    assert detect_module([]) is None


def test_detect_module_file_directly_under_tests_returns_none():
    assert detect_module(["projects/AcuHMI_1_7/tests/__init__.py"]) is None


def test_update_latest_writes_pointer(tmp_path: Path):
    run_dir = tmp_path / "AcuHMI_1_7" / "20260610_153000"
    run_dir.mkdir(parents=True)
    update_latest(run_dir, reports_root=tmp_path)
    latest = tmp_path / "latest"
    # 软链或退化的 latest.txt 至少有一个指向 run_dir
    if latest.exists():
        assert Path(latest).resolve() == run_dir.resolve()
    else:
        assert (tmp_path / "latest.txt").read_text(encoding="utf-8").strip() == str(run_dir)
