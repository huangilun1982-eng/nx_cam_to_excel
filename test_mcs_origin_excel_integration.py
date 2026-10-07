# -*- coding: utf-8 -*-
"""
整合單元測試：test_mcs_origin_excel_integration.py
驗證 MCS 加工原點標註寫入 Excel 工單示圖下方 (方案 B: A37 儲存格)
"""
import os
import sys
import unittest
import tempfile
import datetime
import subprocess
from unittest.mock import MagicMock

# 模擬 NXOpen 供外部環境測試
for mod in ["NXOpen", "NXOpen.CAM", "NXOpen.UF", "NXOpen.Gateway"]:
    if mod not in sys.modules:
        sys.modules[mod] = MagicMock()

from nx_cam_to_excel import export_multipage_via_vbs

class DummyPart:
    Leaf = "TEST_MCS_ORIGIN_PART"
    FullPath = r"C:\NX_Standard\Template\TEST_MCS_ORIGIN.prt"

class TestMCSOriginExcelIntegration(unittest.TestCase):
    def setUp(self):
        self.template_path = r"c:\NX_Standard\Template\ShopDoc_Template.xlsx"
        self.out_xlsx = os.path.join(tempfile.gettempdir(), f"test_mcs_integ_{datetime.datetime.now().strftime('%H%M%S%f')}.xlsx")

    def tearDown(self):
        if os.path.exists(self.out_xlsx):
            try:
                os.remove(self.out_xlsx)
            except Exception:
                pass

    def test_mcs_origin_in_cell_a37(self):
        self.assertTrue(os.path.exists(self.template_path), "工單範本不存在")

        pages = [{
            "stage": "TOP",
            "rows": [
                {
                    "type": "op",
                    "data": {
                        "seq": 1,
                        "op_name": "OP01_ROUGH",
                        "tool_number": "T01",
                        "tool_spec_display": "E10",
                        "flute_length": "25.0",
                        "holder_length": "30.0",
                        "time": "00:05:00",
                        "note": "S:5000 F:2000"
                    }
                }
            ],
            "show_image": False,
            "image_top_row": 18,
            "image_bottom_row": 36,
            "is_appendix_image_page": False
        }]

        header_info = {
            "drawing_number": "DWG-12345678901234",
            "drawing_name": "TEST_MCS_ORIGIN_PART",
            "blank_size": "150x100x35",
            "part_number": "P9999",
            "holes": "8"
        }

        stage_origins = {
            "TOP": "加工原點【G54】：X=MID, Y=MID, Z=TOP"
        }

        export_multipage_via_vbs(
            pages=pages,
            template_path=self.template_path,
            output_path=self.out_xlsx,
            work_part=DummyPart(),
            header_info=header_info,
            stage_images={},
            stage_origins=stage_origins
        )

        self.assertTrue(os.path.exists(self.out_xlsx), "Excel 輸出檔案應存在")

        verify_vbs = f"""
Dim objExcel, objWb, ws
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objWb = objExcel.Workbooks.Open("{os.path.abspath(self.out_xlsx)}")
Set ws = objWb.Sheets(1)

a37_val = ws.Range("A37").Value
a37_bold = ws.Range("A37").Font.Bold
a37_size = ws.Range("A37").Font.Size
j37_val = ws.Range("J37").Value

objWb.Close False
objExcel.Quit

WScript.Echo "A37_VAL=" & a37_val
WScript.Echo "A37_BOLD=" & a37_bold
WScript.Echo "A37_SIZE=" & a37_size
WScript.Echo "J37_VAL=" & j37_val
"""
        tmp_check_vbs = os.path.join(tempfile.gettempdir(), f"check_mcs_{datetime.datetime.now().strftime('%H%M%S%f')}.vbs")
        with open(tmp_check_vbs, "w", encoding="cp950") as f:
            f.write(verify_vbs)

        res = subprocess.run(["cscript.exe", "//Nologo", tmp_check_vbs], capture_output=True, text=True)
        if os.path.exists(tmp_check_vbs):
            os.remove(tmp_check_vbs)

        out_text = res.stdout
        self.assertIn("A37_VAL=加工原點【G54】：X=MID, Y=MID, Z=TOP", out_text)
        self.assertIn("A37_BOLD=True", out_text)
        self.assertIn("A37_SIZE=10", out_text)
        self.assertIn("J37_VAL=1-1", out_text)

if __name__ == "__main__":
    unittest.main()
