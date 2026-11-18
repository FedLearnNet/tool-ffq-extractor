import pytest


pyfed = pytest.importorskip("pyfedappwrap")


def test_config_dataclasses_importable():
    # Import from repo root modules used by main.py / app.py
    import sys
    from pathlib import Path

    root = Path(__file__).resolve().parents[1]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

    from config import MyAppConfig, MyAppInputConfig, MyAppOutputConfig

    cfg = MyAppConfig(
        workers=2,
        sheet="",
        date_cell="A1",
        patient_id_cell="A3",
    )
    assert cfg.workers == 2
    assert cfg.sheet == ""
    assert cfg.date_cell == "A1"
    assert cfg.patient_id_cell == "A3"
    assert "strict" not in MyAppConfig.__dataclass_fields__
    assert MyAppInputConfig.__dataclass_fields__["ffq_archive"]
    assert MyAppOutputConfig.__dataclass_fields__["nutrition_csv"]
