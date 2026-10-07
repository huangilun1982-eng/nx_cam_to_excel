# -*- coding: utf-8 -*-
"""
整合單元測試：test_tool_conflict_excel_integration.py
驗證刀號重複性衝突檢驗模組與 Excel 工單匯出端 (VBScript / COM) 的無縫整合：
1. 衝突檢驗精準標記 has_tool_conflict 與 tool_conflict_msg。
2. 正常工步不產生任何干擾與標記。
3. 匯出 Excel 後，衝突列刀號儲存格正確套用淡紅底色 (13421823) 與深紅粗體字 (180)，備註正確追加衝突警示。
"""
import os
import sys
import unittest
import tempfile
import datetime
import subprocess
from unittest.mock import MagicMock

# 模擬 NXOpen 供外部環境執行測試
for mod in ["NXOpen", "NXOpen.CAM", "NXOpen.UF", "NXOpen.Gateway"]:
    if mod not in sys.modules:
        sys.modules[mod] = MagicMock()

from nx_cam_to_excel import export_multipage_via_vbs
from tool_conflict_detector import detect_tool_conflicts_for_operations

class DummyPart:
    Leaf = "TEST_INTEGRATION_PART"
    FullPath = r"C:\NX_Standard\Template\TEST_INTEGRATION.prt"

class TestToolConflictExcelIntegration(unittest.TestCase):
    def setUp(self):
        self.template_path = r"c:\NX_Standard\Template\ShopDoc_Template.xlsx"
        self.out_xlsx = os.path.join(tempfile.gettempdir(), f"test_conflict_integ_{datetime.datetime.now().strftime('%H%M%S%f')}.xlsx")

    def tearDown(self):
        if os.path.exists(self.out_xlsx):
            try:
                os.remove(self.out_xlsx)
            except Exception:
                pass

    def test_e2e_tool_conflict_excel_highlighting(self):
        self.assertTrue(os.path.exists(self.template_path), "工單母版不存在")

        ops = [
            {
                "seq": 1,
                "op_name": "ROUGH_OP",
                "tool_number": "T01",
                "raw_tool_name": "E10",
                "tool_diameter": "10.0",
                "tool_spec_display": "E10",
                "flute_length": "25.0",
                "holder_length": "30.0",
                "time": "00:05:30",
                "note": "S:5000 F:2000"
            },
            {
                "seq": 2,
                "op_name": "",
                "tool_number": "T02",
                "raw_tool_name": "D6",
                "tool_diameter": "6.0",
                "tool_spec_display": "D6",
                "flute_length": "15.0",
                "holder_length": "20.0",
                "time": "00:03:10",
                "note": "S:6000 F:1500"
            },
            {
                "seq": 3,
                "op_name": "FINISH_OP",
                "tool_number": "T01",
                "raw_tool_name": "E6",
                "tool_diameter": "6.0",
                "tool_spec_display": "E6",
                "flute_length": "18.0",
                "holder_length": "25.0",
                "time": "00:04:20",
                "note": "S:8000 F:1800"
            }
        ]

        conflicts = detect_tool_conflicts_for_operations(ops)
        self.assertEqual(len(conflicts), 1)
        self.assertTrue(ops[0].get("has_tool_conflict"))
        self.assertFalse(ops[1].get("has_tool_conflict", False))
        self.assertTrue(ops[2].get("has_tool_conflict"))

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
            "drawing_name": "TEST_INTEGRATION_PART",
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
            stage_images={}
        )

        self.assertTrue(os.path.exists(self.out_xlsx), "Excel 輸出檔案應成功建立")

        verify_vbs = f"""
Dim objExcel, objWb, ws
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objWb = objExcel.Workbooks.Open("{os.path.abspath(self.out_xlsx)}")
Set ws = objWb.Sheets(1)

c1_color = ws.Cells(7, 3).Interior.Color
c1_font_color = ws.Cells(7, 3).Font.Color
c1_bold = ws.Cells(7, 3).Font.Bold
c1_note = ws.Cells(7, 8).Value

c2_color = ws.Cells(8, 3).Interior.Color
c3_color = ws.Cells(9, 3).Interior.Color

objWb.Close False
objExcel.Quit

WScript.Echo "R7_COLOR=" & c1_color & "|R7_FONT=" & c1_font_color & "|R7_BOLD=" & c1_bold
WScript.Echo "R8_COLOR=" & c2_color
WScript.Echo "R9_COLOR=" & c3_color
WScript.Echo "R7_NOTE=" & c1_note
"""
        tmp_check_vbs = os.path.join(tempfile.gettempdir(), f"check_excel_{datetime.datetime.now().strftime('%H%M%S%f')}.vbs")
        with open(tmp_check_vbs, "w", encoding="cp950") as f:
            f.write(verify_vbs)

        res = subprocess.run(["cscript.exe", "//Nologo", tmp_check_vbs], capture_output=True, text=True)
        if os.path.exists(tmp_check_vbs):
            os.remove(tmp_check_vbs)

        out_text = res.stdout
        self.assertIn("R7_COLOR=13421823", out_text)
        self.assertIn("R7_FONT=180", out_text)
        self.assertIn("R7_BOLD=True", out_text)
        self.assertIn("R9_COLOR=13421823", out_text)
        self.assertNotIn("R8_COLOR=13421823", out_text)
        self.assertIn("衝突", out_text)

if __name__ == "__main__":
    unittest.main()
