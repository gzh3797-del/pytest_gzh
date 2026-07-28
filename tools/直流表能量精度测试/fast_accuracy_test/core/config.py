import json

MODELS = {"320", "300", "260"}
CONN_MODES = {"rtu", "tcp"}


def load_config(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    if cfg.get("conn_mode") not in CONN_MODES:
        raise ValueError(f"conn_mode must be one of {CONN_MODES}, got {cfg.get('conn_mode')!r}")
    if cfg.get("device_model") not in MODELS:
        raise ValueError(f"device_model must be one of {MODELS}, got {cfg.get('device_model')!r}")

    cfg["is_dual"] = (cfg["device_model"] == "260")
    cfg.setdefault("word_order", "big")
    cfg.setdefault("input_xlsx", "./test_data/input.xlsx")
    cfg.setdefault("result_dir", "./result")
    return cfg
