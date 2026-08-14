from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]


@pytest.fixture(scope="session")
def xlsx_path() -> Path:
    return REPO_ROOT / "RealExcel.xlsx"


@pytest.fixture(scope="session")
def xlsb_path() -> Path:
    return REPO_ROOT / "RealExcel.xlsb"


@pytest.fixture(scope="session")
def xls_path() -> Path:
    return REPO_ROOT / "tests" / "ExcelReader.Benchmarks" / "Data" / "65K_Records_Data.xls"


@pytest.fixture
def csv_path(tmp_path: Path) -> Path:
    path = tmp_path / "sample.csv"
    path.write_text("name,qty\nwidget,7\ngadget,9\n", encoding="utf-8")
    return path
