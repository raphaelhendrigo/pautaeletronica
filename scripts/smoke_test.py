from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from settings import ConfigError, require_any, require_env
from docx_maker import roman, _format_reinc_relator_label


def _assert(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


def test_roman() -> None:
    _assert(roman(1) == "I", "roman(1) should be I")
    _assert(roman(2) == "II", "roman(2) should be II")
    _assert(roman(4) == "IV", "roman(4) should be IV")
    _assert(roman(9) == "IX", "roman(9) should be IX")
    _assert(roman(20) == "XX", "roman(20) should be XX")


def test_reinc_relator_label() -> None:
    label = _format_reinc_relator_label("Fulano de Tal", "CONSELHEIRO")
    _assert(label == "RELATOR CONSELHEIRO FULANO DE TAL", "reinclusao relator label should not be numbered")


def test_config_missing() -> None:
    missing_ok = False
    try:
        require_env("SMOKE_MISSING_ENV")
    except ConfigError:
        missing_ok = True
    _assert(missing_ok, "require_env should fail for missing variable")

    missing_any_ok = False
    try:
        require_any(["SMOKE_MISSING_ENV_A", "SMOKE_MISSING_ENV_B"])
    except ConfigError:
        missing_any_ok = True
    _assert(missing_any_ok, "require_any should fail when all are missing")


def main() -> None:
    test_roman()
    test_reinc_relator_label()
    test_config_missing()
    print("smoke_test: ok")


if __name__ == "__main__":
    main()
