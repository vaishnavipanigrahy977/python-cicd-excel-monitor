import os
from unittest.mock import MagicMock, patch
import openpyxl
from monitor import check_runtime_health, update_excel_report


@patch("monitor.requests.get")
def test_runtime_check_passed(mock_get):
    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_get.return_value = mock_resp

    status, code, latency = check_runtime_health("https://example.com")
    assert status == "PASSED"
    assert code == 200
    assert latency >= 0


def test_excel_generation(tmp_path, monkeypatch):
    test_excel = tmp_path / "test_report.xlsx"
    monkeypatch.setattr("monitor.EXCEL_FILE", str(test_excel))
    update_excel_report("PASSED", 200, 125.5)

    assert os.path.exists(test_excel)
    wb = openpyxl.load_workbook(test_excel)
    ws = wb.active
    assert ws.cell(row=2, column=3).value == "PASSED"
    
