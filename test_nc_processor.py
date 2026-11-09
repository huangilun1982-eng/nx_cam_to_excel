# -*- coding: utf-8 -*-
"""
單元測試：test_nc_processor.py
驗證 NC 文字清洗、標準註解檔頭注入、臥式機型 12 行專屬檔頭置換與安全落盤
"""
import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock

for mod in ["NXOpen", "NXOpen.CAM", "NXOpen.UF", "NXOpen.Gateway"]:
    if mod not in sys.modules:
        sys.modules[mod] = MagicMock()

from nc_processor import (
    sanitize_nc_text,
    build_nc_header_comment_block,
    transform_nc_content_by_profile,
    parse_nc_initial_params,
    process_and_export_nc_tasks,
    safe_write_nc_file
)
from machine_profile_manager import (
    load_machine_profiles,
    get_last_selected_profile_id,
    save_last_selected_profile_id,
    get_post_settings,
    save_post_settings
)
from nx_cam_to_excel import merge_consecutive_tools

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

    def test_post_settings_persistence(self):
        """
        驗證機台後處理器設定完整持久化：包含機型、後處理名稱與自訂路徑
        """
        # 保存前狀態讀取
        orig_settings = get_post_settings()

        try:
            # 1. 寫入完整自訂設定
            save_post_settings(
                profile_id="RAW",
                post_name="Mazak_AC",
                custom_post_path="C:/custom/my_post.tcl"
            )
            settings = get_post_settings()
            self.assertEqual(settings["last_selected_id"], "RAW")
            self.assertEqual(settings["last_selected_post"], "Mazak_AC")
            self.assertEqual(settings["last_custom_post_path"], "C:/custom/my_post.tcl")

            # 2. 部分更新 (僅更新後處理器名稱，保留其他欄位)
            save_post_settings(post_name="YCM_noCheck")
            settings2 = get_post_settings()
            self.assertEqual(settings2["last_selected_id"], "RAW")
            self.assertEqual(settings2["last_selected_post"], "YCM_noCheck")
            self.assertEqual(settings2["last_custom_post_path"], "C:/custom/my_post.tcl")

            # 3. 驗證 get_last_selected_profile_id 同步相容
            self.assertEqual(get_last_selected_profile_id(), "RAW")
        finally:
            # 復原原始設定
            save_post_settings(
                profile_id=orig_settings.get("last_selected_id", "FANUC_HORIZONTAL"),
                post_name=orig_settings.get("last_selected_post", "Fanuc_2026"),
                custom_post_path=orig_settings.get("last_custom_post_path", "")
            )

    def test_nc_post_dialog_settings_restoration(self):
        """
        驗證 nc_post_dialog 視窗能自動自持久化設定恢復上次選擇的機台與後處理器
        """
        import tkinter as tk
        from nc_post_dialog import NcPostDialogApp

        orig_settings = get_post_settings()

        try:
            # 設定模擬上次選擇：BROTHER + YCM_noCheck
            save_post_settings(
                profile_id="BROTHER",
                post_name="YCM_noCheck",
                custom_post_path=""
            )

            root = tk.Tk()
            root.withdraw()
            try:
                # 不傳遞 default_post 與 default_machine_profile，驗證自動載入持久化值
                app = NcPostDialogApp(root, config={"nc_tasks": []})
                self.assertEqual(app.combo_post.get(), "YCM_noCheck")
                self.assertEqual(app.profile_var.get(), "BROTHER 機台 (攻牙機/小型加工中心)")

                # 測試自動觸發 export_only，應再次更新並記住當前設定
                app.on_export_only()
                self.assertEqual(app.result["action"], "export_only")
                self.assertEqual(app.result["machine_profile"], "BROTHER")
                self.assertEqual(app.result["postprocessor_name"], "YCM_noCheck")
            finally:
                try:
                    root.destroy()
                except Exception:
                    pass
        finally:
            save_post_settings(
                profile_id=orig_settings.get("last_selected_id", "FANUC_HORIZONTAL"),
                post_name=orig_settings.get("last_selected_post", "Fanuc_2026"),
                custom_post_path=orig_settings.get("last_custom_post_path", "")
            )

    def test_fanuc_horizontal_dynamic_coordinates_and_wcs(self):
        """
        驗證 Fanuc 臥式機型：動態解析並保留真實起刀點 X/Y 座標與工段 WCS (如 G55)，絕不覆蓋為 X0. Y0. 或死鎖 G54
        """
        raw_nc_g55 = """%
O1002
G91 G28 Z0.
T03 M06
S6500 M03
G55 G00 X155.432 Y-45.678
G43 H03 Z30.
M08
G01 Z-5. F800.
X200. Y100.
M30
%"""
        header_info = {
            "drawing_name": "PART_G55_DYNAMIC",
            "drawing_number": "DWG-5555",
            "blank_size": "150x150x50"
        }
        mcs_origin_str = "加工原點【G55】：X=MID, Y=MID, Z=TOP"
        raw_ops = [
            {
                "tool_number": "T03",
                "note": "S:6500 F:800",
                "clean_op_name": "OP20_FINISH",
                "tool_spec_display": "D8R0"
            }
        ]

        result_nc = transform_nc_content_by_profile(
            original_nc=raw_nc_g55,
            profile_id="FANUC_HORIZONTAL",
            header_info=header_info,
            mcs_origin_str=mcs_origin_str,
            raw_ops=raw_ops,
            program_name="OP20_G55"
        )

        # 驗證 B0 旋轉使用動態工段 WCS (G55)
        self.assertIn("G90G0G55B0.", result_nc)
        # 驗證起刀點座標使用真實解析值 (X155.432 Y-45.678)，絕非死鎖之 X0. Y0.
        self.assertIn("G90 G55 G00 X155.432 Y-45.678", result_nc)
        # 驗證換刀、轉速、安全 Z 與冷卻液正常置換
        self.assertIn("M06 T3", result_nc)
        self.assertIn("S6500 M03", result_nc)
        self.assertIn("G0 G43 H3 Z30.", result_nc)
        self.assertIn("M08", result_nc)
        # 驗證後續切削本體路徑 100% 完整無缺
        self.assertIn("G01 Z-5. F800.", result_nc)
        self.assertIn("X200. Y100.", result_nc)
        self.assertIn("M30", result_nc)

    def test_all_operations_preserved_in_batch_postprocess(self):
        """
        驗證批次後處理同刀號工序完整保留：
        1. 即使 Excel 工單因同刀號合併為一列 (merge_consecutive_tools)，NC 後處理仍能提取全部工序物件
        2. 直接傳入 nc_tasks 亦能完整後處理所有工序物件
        """
        # 模擬群組內 3 道使用相同刀號 T01 之工序
        op1 = {'tool_number': '1', 'raw_tool_name': 'D10', 'tool_diameter': '10.0', 'flute_length': '20.0', 'holder_length': '40.0', '_nx_op_obj': 'OP_OBJ_1', 'clean_op_name': 'OP_ROUGH'}
        op2 = {'tool_number': '1', 'raw_tool_name': 'D10', 'tool_diameter': '10.0', 'flute_length': '20.0', 'holder_length': '40.0', '_nx_op_obj': 'OP_OBJ_2', 'clean_op_name': 'OP_SEMIFINISH'}
        op3 = {'tool_number': '1', 'raw_tool_name': 'D10', 'tool_diameter': '10.0', 'flute_length': '20.0', 'holder_length': '40.0', '_nx_op_obj': 'OP_OBJ_3', 'clean_op_name': 'OP_FINISH'}

        # 經 Excel 刀具合併後，工步列數為 1
        merged = merge_consecutive_tools([op1, op2, op3])
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0]["_merged_nx_ops"], ['OP_OBJ_1', 'OP_OBJ_2', 'OP_OBJ_3'])

        processed_chunks = [{
            'group_name': 'MAIN',
            'parent_program': 'PROG_MERGED_TEST',
            'stage': '通用工段',
            'operations': merged
        }]

        captured_ops_from_chunks = []
        def mock_postprocess(cam_objects, post_name, output_file, units):
            captured_ops_from_chunks.extend(cam_objects)
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write('%\nO1001\nM08\nM30\n%\n')

        mock_cam_setup = MagicMock()
        mock_cam_setup.Postprocess.side_effect = mock_postprocess

        tmp_out = tempfile.mkdtemp()
        succ, fail = process_and_export_nc_tasks(
            mock_cam_setup, processed_chunks, 'FANUC_HORIZONTAL', {}, {}, tmp_out
        )
        self.assertEqual(succ, 1)
        self.assertEqual(fail, 0)
        # 核心驗證：3 道工序全部被傳送至 cam_setup.Postprocess，無任何遺漏！
        self.assertEqual(captured_ops_from_chunks, ['OP_OBJ_1', 'OP_OBJ_2', 'OP_OBJ_3'])

        # 驗證模式 2：傳入 nc_tasks
        captured_ops_from_nc_tasks = []
        def mock_postprocess_nc_tasks(cam_objects, post_name, output_file, units):
            captured_ops_from_nc_tasks.extend(cam_objects)
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write('%\nO1002\nM08\nM30\n%\n')

        mock_cam_setup.Postprocess.side_effect = mock_postprocess_nc_tasks
        nc_tasks = [{
            'program_name': 'PROG_NC_TASKS_TEST',
            'operations': ['OP_OBJ_A', 'OP_OBJ_B', 'OP_OBJ_C'],
            'op_names': ['OP_A', 'OP_B', 'OP_C'],
            'stage': '通用工段'
        }]

        succ2, fail2 = process_and_export_nc_tasks(
            mock_cam_setup, [], 'FANUC_HORIZONTAL', {}, {}, tmp_out, nc_tasks=nc_tasks
        )
        self.assertEqual(succ2, 1)
        self.assertEqual(fail2, 0)
        self.assertEqual(captured_ops_from_nc_tasks, ['OP_OBJ_A', 'OP_OBJ_B', 'OP_OBJ_C'])

if __name__ == "__main__":
    unittest.main()
