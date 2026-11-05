# -*- coding: utf-8 -*-
"""
單元測試：test_sequence_style_integration.py
驗證 export_multipage_via_vbs 在各種序號風格下的 VBS 生成行為：
1. per_page: 每頁獨立起算 (1, 2, 3...)
2. stage_prefix: 工段前綴 (M1-01, M1-02, M2-01...)
3. global_continuous: 全工單連續累加
4. zero_padded: 兩位補零 (01, 02...)
5. step_10: 工藝步長 (10, 20...)
"""

import unittest
import os
import sys
from unittest.mock import MagicMock, patch

# 模擬 NXOpen 供外部環境執行測試
for mod in ["NXOpen", "NXOpen.CAM", "NXOpen.UF", "NXOpen.Gateway"]:
    if mod not in sys.modules:
        sys.modules[mod] = MagicMock()

from nx_cam_to_excel import export_multipage_via_vbs


class DummyPart:
    def __init__(self, name="PART_TEST", full_path=r"C:\Test\PART_TEST.prt"):
        self.Leaf = name
        self.FullPath = full_path


class TestSequenceStyleIntegration(unittest.TestCase):

    def setUp(self):
        self.dummy_part = DummyPart()
        self.header_info = {
            "drawing_name": "PART_TEST",
            "drawing_number": "12345678901234",
            "blank_size": "100x100x20",
            "part_number": "PO-999",
            "holes": "4",
            "thickness": "20"
        }
        self.pages_data = [
            {
                "stage": "M1",
                "show_image": True,
                "image_top_row": 18,
                "image_bottom_row": 36,
                "is_appendix_image_page": False,
                "rows": [
                    {"type": "op", "data": {"tool_number": "T01", "tool_spec_display": "D10", "op_name": "OP1"}},
                    {"type": "op", "data": {"tool_number": "T02", "tool_spec_display": "D8", "op_name": "OP2"}}
                ]
            },
            {
                "stage": "M2",
                "show_image": True,
                "image_top_row": 18,
                "image_bottom_row": 36,
                "is_appendix_image_page": False,
                "rows": [
                    {"type": "op", "data": {"tool_number": "T03", "tool_spec_display": "D6", "op_name": "OP3"}}
                ]
            }
        ]

    @patch("subprocess.run")
    def test_per_page_sequence_vbs(self, mock_subproc):
        mock_subproc.return_value = MagicMock(returncode=0)
        template_file = r"C:\NX_Standard\Template\ShopDoc_Template.xlsx"

        captured_vbs = []
        def fake_open(path, mode="r", *args, **kwargs):
            if "w" in mode and path.endswith(".vbs"):
                class FakeFile:
                    def write(self, s):
                        captured_vbs.append(s)
                    def __enter__(self): return self
                    def __exit__(self, *a): pass
                return FakeFile()
            return orig_open(path, mode, *args, **kwargs)

        orig_open = open
        with patch("builtins.open", side_effect=fake_open):
            export_multipage_via_vbs(
                self.pages_data, template_file, r"C:\Test\out.xlsx", self.dummy_part,
                self.header_info, sequence_style="per_page"
            )

        vbs_content = "".join(captured_vbs)
        # 第一頁：1, 2
        # 第二頁：1
        self.assertIn('ws.Cells(7, 1).Value = "1"', vbs_content)
        self.assertIn('ws.Cells(8, 1).Value = "2"', vbs_content)

    @patch("subprocess.run")
    def test_stage_prefix_sequence_vbs(self, mock_subproc):
        mock_subproc.return_value = MagicMock(returncode=0)
        template_file = r"C:\NX_Standard\Template\ShopDoc_Template.xlsx"

        captured_vbs = []
        def fake_open(path, mode="r", *args, **kwargs):
            if "w" in mode and path.endswith(".vbs"):
                class FakeFile:
                    def write(self, s):
                        captured_vbs.append(s)
                    def __enter__(self): return self
                    def __exit__(self, *a): pass
                return FakeFile()
            return orig_open(path, mode, *args, **kwargs)

        orig_open = open
        with patch("builtins.open", side_effect=fake_open):
            export_multipage_via_vbs(
                self.pages_data, template_file, r"C:\Test\out.xlsx", self.dummy_part,
                self.header_info, sequence_style="stage_prefix"
            )

        vbs_content = "".join(captured_vbs)
        # 第一頁 (M1)：M1-01, M1-02
        # 第二頁 (M2)：M2-01
        self.assertIn('ws.Cells(7, 1).Value = "M1-01"', vbs_content)
        self.assertIn('ws.Cells(8, 1).Value = "M1-02"', vbs_content)
        self.assertIn('ws.Cells(7, 1).Value = "M2-01"', vbs_content)

    @patch("subprocess.run")
    def test_global_continuous_sequence_vbs(self, mock_subproc):
        mock_subproc.return_value = MagicMock(returncode=0)
        template_file = r"C:\NX_Standard\Template\ShopDoc_Template.xlsx"

        captured_vbs = []
        def fake_open(path, mode="r", *args, **kwargs):
            if "w" in mode and path.endswith(".vbs"):
                class FakeFile:
                    def write(self, s):
                        captured_vbs.append(s)
                    def __enter__(self): return self
                    def __exit__(self, *a): pass
                return FakeFile()
            return orig_open(path, mode, *args, **kwargs)

        orig_open = open
        with patch("builtins.open", side_effect=fake_open):
            export_multipage_via_vbs(
                self.pages_data, template_file, r"C:\Test\out.xlsx", self.dummy_part,
                self.header_info, sequence_style="global_continuous"
            )

        vbs_content = "".join(captured_vbs)
        # 第一頁：1, 2
        # 第二頁：3
        self.assertIn('ws.Cells(7, 1).Value = "1"', vbs_content)
        self.assertIn('ws.Cells(8, 1).Value = "2"', vbs_content)
        self.assertIn('ws.Cells(7, 1).Value = "3"', vbs_content)

    @patch("subprocess.run")
    def test_custom_seq_label_vbs(self, mock_subproc):
        """測試自訂工段文字標記直接寫入 Excel 儲存格，完全依照輸入文字顯示，不加流水序號"""
        mock_subproc.return_value = MagicMock(returncode=0)
        template_file = r"C:\NX_Standard\Template\ShopDoc_Template.xlsx"

        pages_with_custom = [
            {
                "stage": "M1",
                "show_image": True,
                "image_top_row": 18,
                "image_bottom_row": 36,
                "is_appendix_image_page": False,
                "rows": [
                    {"type": "op", "data": {"tool_number": "T01", "tool_spec_display": "D10", "op_name": "OP1", "custom_seq": "OP10"}},
                    {"type": "op", "data": {"tool_number": "T02", "tool_spec_display": "D8", "op_name": "OP2", "custom_seq": "OP10"}}
                ]
            }
        ]

        captured_vbs = []
        def fake_open(path, mode="r", *args, **kwargs):
            if "w" in mode and path.endswith(".vbs"):
                class FakeFile:
                    def write(self, s):
                        captured_vbs.append(s)
                    def __enter__(self): return self
                    def __exit__(self, *a): pass
                return FakeFile()
            return orig_open(path, mode, *args, **kwargs)

        orig_open = open
        with patch("builtins.open", side_effect=fake_open):
            export_multipage_via_vbs(
                pages_with_custom, template_file, r"C:\Test\out.xlsx", self.dummy_part,
                self.header_info, sequence_style="per_page"
            )

        vbs_content = "".join(captured_vbs)
        # 兩筆資料皆直接填入 "OP10" (不帶 -01, -02 等序號)
        self.assertIn('ws.Cells(7, 1).Value = "OP10"', vbs_content)
        self.assertIn('ws.Cells(8, 1).Value = "OP10"', vbs_content)


    @patch("subprocess.run")
    def test_target_op_separate_screenshot_binding_vbs(self, mock_subproc):
        """
        用戶規則驗證：「工序右方截圖若有勾選 就是表示該工序需要一張截圖」
        驗證當分頁中指定各工序的 target_op 時：
        1. VBS 必須依照 target_op 匹配對應的專屬截圖路徑。
        2. 大圖附圖頁的 A37 儲存格必須標記工序專屬加工示圖名稱。
        """
        mock_subproc.return_value = MagicMock(returncode=0)
        template_file = r"C:\NX_Standard\Template\ShopDoc_Template.xlsx"

        pages = [
            # 第 1 頁：刀具頁
            {
                "stage": "通用工段",
                "rows": [
                    {
                        "type": "op",
                        "data": {
                            "program_name": "755613070-PEEK",
                            "op_name": "755613070-PEEK",
                            "custom_seq": "1",
                            "tool_number": "T14",
                            "tool_spec_display": "E4",
                            "time": "00:01:38",
                            "stage": "通用工段"
                        }
                    }
                ],
                "show_image": False,
                "is_appendix_image_page": False
            },
            # 第 2 頁：工序 755613070-PEEK 專屬大圖附圖頁
            {
                "stage": "通用工段",
                "target_op": "755613070-PEEK",
                "rows": [],
                "show_image": True,
                "image_top_row": 7,
                "image_bottom_row": 36,
                "is_appendix_image_page": True
            },
            # 第 3 頁：工序 755613070-A 專屬大圖附圖頁
            {
                "stage": "通用工段",
                "target_op": "755613070-A",
                "rows": [],
                "show_image": True,
                "image_top_row": 7,
                "image_bottom_row": 36,
                "is_appendix_image_page": True
            }
        ]

        dummy_img_peek = r"C:\Test\preview_755613070-PEEK.png"
        dummy_img_a = r"C:\Test\preview_755613070-A.png"

        stage_images = {
            "755613070-PEEK": dummy_img_peek,
            "755613070-A": dummy_img_a
        }

        captured_vbs = []
        def fake_open(path, mode="r", *args, **kwargs):
            if "w" in mode and path.endswith(".vbs"):
                class FakeFile:
                    def write(self, s):
                        captured_vbs.append(s)
                    def __enter__(self): return self
                    def __exit__(self, *a): pass
                return FakeFile()
            return orig_open(path, mode, *args, **kwargs)

        orig_open = open
        with patch("os.path.exists", return_value=True), patch("builtins.open", side_effect=fake_open):
            export_multipage_via_vbs(
                pages,
                template_file,
                r"C:\Test\out.xlsx",
                self.dummy_part,
                self.header_info,
                stage_images=stage_images,
                stage_origins={"通用工段": "G54 MID TOP"}
            )

        vbs_content = "".join(captured_vbs)

        # 斷言 1：VBS 中包含 PEEK 專屬截圖路徑與工序專屬標籤
        self.assertIn("preview_755613070-PEEK.png", vbs_content)
        self.assertIn("工序【755613070-PEEK】專屬加工示圖", vbs_content)

        # 斷言 2：VBS 中包含 A 專屬截圖路徑與工序專屬標籤
        self.assertIn("preview_755613070-A.png", vbs_content)
        self.assertIn("工序【755613070-A】專屬加工示圖", vbs_content)


    @patch("subprocess.run")
    def test_grid_image_layout_vbs(self, mock_subproc):
        """
        驗證模式 B 多圖網格排列 (最多 6 張圖) 的 VBS 生成：
        當單頁有 4 張圖時：
        1. VBS 必須生成 2x2 網格計算邏輯 (cellW, cellH, cL0~cL3, cT0~cT3)
        2. 必須呼叫 4 次 Shapes.AddPicture 插入 4 張圖
        3. A37 儲存格必須呈現完整 4 張圖標註：【圖1: OP10】  【圖2: OP20】  【圖3: OP30】  【圖4: OP40】
        """
        mock_subproc.return_value = MagicMock(returncode=0)
        template_file = r"C:\NX_Standard\Template\ShopDoc_Template.xlsx"

        pages = [
            {
                "stage": "通用工段",
                "rows": [
                    {"type": "op", "data": {"tool_number": "T01", "tool_spec_display": "D10", "op_name": "OP10"}}
                ],
                "show_image": False,
                "is_appendix_image_page": False
            },
            {
                "stage": "通用工段",
                "target_ops": ["OP10", "OP20", "OP30", "OP40"],
                "target_op": "OP10",
                "rows": [],
                "show_image": True,
                "image_top_row": 7,
                "image_bottom_row": 36,
                "is_appendix_image_page": True,
                "image_layout": "grid"
            }
        ]

        stage_images = {
            "OP10": r"C:\Test\img_op10.png",
            "OP20": r"C:\Test\img_op20.png",
            "OP30": r"C:\Test\img_op30.png",
            "OP40": r"C:\Test\img_op40.png"
        }

        captured_vbs = []
        def fake_open(path, mode="r", *args, **kwargs):
            if "w" in mode and path.endswith(".vbs"):
                class FakeFile:
                    def write(self, s):
                        captured_vbs.append(s)
                    def __enter__(self): return self
                    def __exit__(self, *a): pass
                return FakeFile()
            return orig_open(path, mode, *args, **kwargs)

        orig_open = open
        with patch("os.path.exists", return_value=True), patch("builtins.open", side_effect=fake_open):
            export_multipage_via_vbs(
                pages,
                template_file,
                r"C:\Test\out.xlsx",
                self.dummy_part,
                self.header_info,
                stage_images=stage_images,
                stage_origins={"通用工段": "G54 MID TOP"}
            )

        vbs_content = "".join(captured_vbs)

        # 斷言 1：VBS 中包含 4 張圖片路徑與 AddPicture 調用
        self.assertIn("img_op10.png", vbs_content)
        self.assertIn("img_op20.png", vbs_content)
        self.assertIn("img_op30.png", vbs_content)
        self.assertIn("img_op40.png", vbs_content)
        self.assertIn("Set shp0 = ws.Shapes.AddPicture", vbs_content)
        self.assertIn("Set shp1 = ws.Shapes.AddPicture", vbs_content)
        self.assertIn("Set shp2 = ws.Shapes.AddPicture", vbs_content)
        self.assertIn("Set shp3 = ws.Shapes.AddPicture", vbs_content)

        # 斷言 2：VBS 中包含網格動態座標計算
        self.assertIn("cellW = (boxW - gapX) / 2", vbs_content)
        self.assertIn("cellH = (boxH - (1) * gapY) / 2", vbs_content)

        # 斷言 3：A37 包含完整多圖標註
        self.assertIn("【圖1: OP10】  【圖2: OP20】  【圖3: OP30】  【圖4: OP40】", vbs_content)

if __name__ == "__main__":
    unittest.main()
