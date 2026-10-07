# -*- coding: utf-8 -*-
"""
整合單元測試：test_tool_safety_excel_integration.py
驗證刀具有效刃長與懸伸安全評估融入 Excel 備註欄
"""
import os
import sys
import unittest
import tempfile
import datetime
import subprocess
from unittest.mock import MagicMock

# 模擬 NXOpen
for mod in ["NXOpen", "NXOpen.CAM", "NXOpen.UF", "NXOpen.Gateway"]:
    if mod not in sys.modules:
        sys.modules[mod] = MagicMock()

from nx_cam_to_excel import export_multipage_via_vbs
from tool_safety_evaluator import evaluate_operations_tool_safety

class DummyPart:
    Leaf = "TEST_TOOL_SAFETY_PART"
    FullPath = r"C:\NX_Standard\Template\TEST_TOOL_SAFETY.prt"

class TestToolSafetyExcelIntegration(unittest.TestCase):
    def setUp(self):
        self.template_path = r"c:\NX_Standard\Template\ShopDoc_Template.xlsx"
        self.out_xlsx = os.path.join(tempfile.gettempdir(), f"test_safety_integ_{datetime.datetime.now().strftime('%H%M%S%f')}.xlsx")

    def tearDown(self):
        if os.path.exists(self.out_xlsx):
            try:
                os.remove(self.out_xlsx)
            except Exception:
                pass

    def test_tool_safety_warning_in_excel_note(self):
        self.assertTrue(os.path.exists(self.template_path), "工單母版不存在")

        ops = [
            {
                "seq": 1,
                "op_name": "DEEP_POCKET",
                "tool_number": "T01",
                "tool_spec_display": "E10",
                "flute_length": "20.0",
                "holder_length": "35.0",
                "time": "00:08:00",
                "note": "S:5000 F:2000",
                "tool_safety_warning": "【注意: 切深28.0 > 刃長20.0】"
            },
            {
                "seq": 2,
                "op_name": "NORMAL_SURFACE",
                "tool_number": "T02",
                "tool_spec_display": "D6",
                "flute_length": "25.0",
                "holder_length": "30.0",
                "time": "00:03:00",
                "note": "S:6000 F:1500"
            }
        ]

        pages = [{
            "stage": "TOP",
            "rows": [{"type": "op", "data": op} for op in ops],
            "show_image": False,
            "image_top_row": 18,
            "image_bottom_row": 36,
            "is_appendix_image_page": False
        }]

        header_info = {
            "drawing_number": "DWG-12345678901234",
            "drawing_name": "TEST_TOOL_SAFETY_PART",
            "blank_size": "100x100x50",
            "part_number": "P12345",
            "holes": "4"
        }

        export_multipage_via_vbs(
            pages=pages,
            template_path=self.template_path,
            output_path=self.out_xlsx,
            work_part=DummyPart(),
            header_info=header_info,
            stage_images={},
            stage_origins={"TOP": "加工原點【G54】：X=MID, Y=MID, Z=TOP"}
        )

        self.assertTrue(os.path.exists(self.out_xlsx), "Excel 輸出檔案應存在")

        verify_vbs = f"""
Dim objExcel, objWb, ws
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objWb = objExcel.Workbooks.Open("{os.path.abspath(self.out_xlsx)}")
Set ws = objWb.Sheets(1)

r7_note = ws.Cells(7, 8).Value
r8_note = ws.Cells(8, 8).Value

objWb.Close False
objExcel.Quit

WScript.Echo "R7_NOTE=" & r7_note
WScript.Echo "R8_NOTE=" & r8_note
"""
        tmp_check_vbs = os.path.join(tempfile.gettempdir(), f"check_safety_{datetime.datetime.now().strftime('%H%M%S%f')}.vbs")
        with open(tmp_check_vbs, "w", encoding="cp950") as f:
            f.write(verify_vbs)

        res = subprocess.run(["cscript.exe", "//Nologo", tmp_check_vbs], capture_output=True, text=True)
        if os.path.exists(tmp_check_vbs):
            os.remove(tmp_check_vbs)

        out_text = res.stdout
        self.assertIn("切深28.0 > 刃長20.0", out_text)
        self.assertIn("S:5000 F:2000", out_text)
        self.assertNotIn("切深", out_text.splitlines()[1])

if __name__ == "__main__":
    unittest.main()
