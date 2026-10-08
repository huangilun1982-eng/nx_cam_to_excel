# -*- coding: utf-8 -*-
"""
單元測試：test_nc_processor.py
驗證 NC 文字清洗、標準註解檔頭注入、臥式機型 12 行專屬檔頭置換與安全落盤
"""
import os
import sys
import tempfile
import unittest

from nc_processor import (
    sanitize_nc_text,
    build_nc_header_comment_block,
    transform_nc_content_by_profile,
    safe_write_nc_file
)
from machine_profile_manager import (
    load_machine_profiles,
    get_last_selected_profile_id,
    save_last_selected_profile_id
)

class TestNCProcessor(unittest.TestCase):
    def test_sanitize_nc_text(self):
        raw = "（加工說明）：直徑＝10.0，原點【G54】"
        sanitized = sanitize_nc_text(raw)
        self.assertNotIn("（", sanitized)
        self.assertNotIn("）", sanitized)
        self.assertNotIn("【", sanitized)
        self.assertNotIn("】", sanitized)
        self.assertNotIn("：", sanitized)
        self.assertIn("(", sanitized)
        self.assertIn(")", sanitized)
        self.assertIn("[", sanitized)
        self.assertIn("]", sanitized)
        self.assertIn(":", sanitized)

    def test_horizontal_header_replacement(self):
        """
        驗證 Fanuc 臥式機型：成功置換使用者指定的 12 行檔頭與 B0/M11/M10/G30
        """
        raw_nc = """%
O1001
(OLD HEADER COMMENT)
G90 G80 G40 G17
T01 M06
S3000 M03
G54 G00 X0. Y0.
G43 H01 Z50.
M08
G00 Z5.
G01 Z-2.0 F500.
X50. Y50.
M09
M05
G91 G28 Z0.
M30
%"""
        header_info = {
            "drawing_name": "TEST_PART_HZ",
            "drawing_number": "DWG-9999",
            "blank_size": "200x150x40"
        }
        mcs_origin_str = "加工原點【G54】：X=MID, Y=MID, Z=TOP"
        raw_ops = [
            {
                "data": {
                    "tool_number": "T10",
                    "note": "S5000 F1500",
                    "tool_spec_display": "D10R0",
                    "op_name": "ROUGH_FACE"
                }
            }
        ]

        result_nc = transform_nc_content_by_profile(
            original_nc=raw_nc,
            profile_id="FANUC_HORIZONTAL",
            header_info=header_info,
            mcs_origin_str=mcs_origin_str,
            raw_ops=raw_ops,
            program_name="OP10_HZ"
        )

        # 驗證包含 12 行臥式標準指令
        self.assertIn("G91G30Z0.", result_nc)
        self.assertIn("G91G30X0.Y0.", result_nc)
        self.assertIn("G90G17G00G40G49G80", result_nc)
        self.assertIn("M11", result_nc)
        self.assertIn("G90G0G54B0.", result_nc)
        self.assertIn("M10", result_nc)
        self.assertIn("M06 T10", result_nc)
        self.assertIn("S5000 M03", result_nc)
        self.assertIn("G17 G40 G80", result_nc)
        self.assertIn("G0 G43 H10 Z50.", result_nc)
        self.assertIn("M08", result_nc)

        # 驗證切削主體指令完整保留
        self.assertIn("G00 Z5.", result_nc)
        self.assertIn("G01 Z-2.0 F500.", result_nc)
        self.assertIn("X50. Y50.", result_nc)
        self.assertIn("M30", result_nc)

        # 驗證舊的註解被替換
        self.assertNotIn("OLD HEADER COMMENT", result_nc)
        self.assertIn("PART NAME    : TEST_PART_HZ", result_nc)

    def test_vertical_and_brother_injection(self):
        """
        驗證 Fanuc 立式與 BROTHER 機台：頂部注入工單資訊，保留原本 G 碼
        """
        raw_nc = """%
O2002
G90 G80 G40 G17
T02 M06
S6000 M03
G54 G00 X10. Y10.
M30
%"""
        header_info = {
            "drawing_name": "BROTHER_PART",
            "drawing_number": "DWG-7777",
            "blank_size": "100x100x20"
        }
        mcs_origin_str = "加工原點【G54】：X=MID, Y=MID, Z=TOP"

        result_nc = transform_nc_content_by_profile(
            original_nc=raw_nc,
            profile_id="BROTHER",
            header_info=header_info,
            mcs_origin_str=mcs_origin_str
        )

        self.assertIn("PART NAME    : BROTHER_PART", result_nc)
        self.assertIn("T02 M06", result_nc)
        self.assertIn("S6000 M03", result_nc)
        self.assertIn("M30", result_nc)

    def test_safe_write_and_backup(self):
        tmp_dir = tempfile.gettempdir()
        test_file = os.path.join(tmp_dir, "test_sample.nc")

        # 首次寫入
        safe_write_nc_file(test_file, "G00 X0. Y0.\nM30\n")
        self.assertTrue(os.path.exists(test_file))

        # 二次寫入，驗證 .bak 備份
        safe_write_nc_file(test_file, "G01 X10. F100.\nM30\n")
        bak_file = test_file + ".bak"
        self.assertTrue(os.path.exists(bak_file))
        with open(bak_file, "r") as f:
            bak_content = f.read()
        self.assertIn("G00 X0. Y0.", bak_content)

        # 清理暫存檔
        try:
            os.remove(test_file)
            os.remove(bak_file)
        except Exception:
            pass

    def test_machine_profile_manager(self):
        profiles = load_machine_profiles()
        self.assertIn("FANUC_HORIZONTAL", profiles)
        self.assertIn("FANUC_VERTICAL", profiles)
        self.assertIn("BROTHER", profiles)

        save_last_selected_profile_id("BROTHER")
        last_id = get_last_selected_profile_id()
        self.assertEqual(last_id, "BROTHER")

        # 復原預設
        save_last_selected_profile_id("FANUC_HORIZONTAL")

if __name__ == "__main__":
    unittest.main()
