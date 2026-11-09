import copy
import json

import pytest

from excel_engine.config import ConfigError, load_config

BASE = {"columns": {"a": {"aliases": ["alpha"], "type": "number"}, "b": {"type": "text"}}}


def make(**over):
    d = copy.deepcopy(BASE)
    d.update(over)
    return load_config(d)


def test_example_config_loads(cfg):
    assert "order_id" in cfg.columns and cfg.computed == {"line_total": "quantity * unit_price"}


def test_unknown_key_rejected():
    with pytest.raises(ConfigError, match="extra|Extra"):
        make(colunms={})


def test_alias_collision_rejected():
    d = copy.deepcopy(BASE)
    d["columns"]["b"]["aliases"] = ["Alpha"]
    with pytest.raises(ConfigError, match="used by both"):
        load_config(d)


def test_bad_strategy_and_type_mismatch():
    d = copy.deepcopy(BASE)
    d["columns"]["b"]["missing"] = "mean"
    with pytest.raises(ConfigError, match="numeric"):
        load_config(d)
    d["columns"]["b"]["missing"] = "bogus"
    with pytest.raises(ConfigError, match="unknown missing strategy"):
        load_config(d)


def test_rule_validation_and_unknown_reference():
    with pytest.raises(ConfigError, match="regex rule needs"):
        make(validation=[{"column": "b", "rule": "regex"}])
    with pytest.raises(ConfigError, match="unknown column 'zzz'"):
        make(validation=[{"column": "zzz", "rule": "required"}])


def test_missing_file_and_bad_json(tmp_path):
    with pytest.raises(ConfigError):
        load_config(tmp_path / "nope.json")
    p = tmp_path / "bad.json"
    p.write_text("{not json")
    with pytest.raises(ConfigError):
        load_config(p)
    p.write_text(json.dumps(BASE))
    assert load_config(p).columns["a"].type == "number"
