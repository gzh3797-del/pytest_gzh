# -*- coding: utf-8 -*-
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../../../.."))

from projects.AcuHMI_1_7.helpers.template_matcher import resolve_template_from_model


def test_known_models_resolve_to_template():
    assert resolve_template_from_model("AcuRev-4110-mA") == "AcuRev4100"
    assert resolve_template_from_model("AcuRev-2100") == "AcuRev2100"
    assert resolve_template_from_model("AcuvimIIW") == "AcuvimIIW"
    assert resolve_template_from_model("AcuvimIIR") == "AcuvimIIR"
    assert resolve_template_from_model("Acuvim3") == "AcuVIM3"


def test_normalization_is_case_and_hyphen_insensitive():
    assert resolve_template_from_model("acurev4110ma") == "AcuRev4100"
    assert resolve_template_from_model("ACUVIM-3") == "AcuVIM3"


def test_unknown_model_returns_none():
    assert resolve_template_from_model("TotallyUnknownMeter") is None
    assert resolve_template_from_model("") is None
