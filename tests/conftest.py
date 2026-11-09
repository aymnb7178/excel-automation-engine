import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from excel_engine.config import load_config  # noqa: E402
from excel_engine.datagen import generate_messy_inputs  # noqa: E402
from excel_engine.pipeline import process  # noqa: E402


@pytest.fixture(scope="session")
def cfg():
    return load_config(ROOT / "examples" / "config.json")


@pytest.fixture(scope="session")
def inputs(tmp_path_factory):
    d = tmp_path_factory.mktemp("messy")
    return d, generate_messy_inputs(d, seed=42)


@pytest.fixture(scope="session")
def result(inputs, cfg):
    return process(inputs[0], cfg)
