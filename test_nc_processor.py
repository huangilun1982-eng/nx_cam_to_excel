# -*- coding: utf-8 -*-
"""
單元測試：test_nc_processor.py
驗證 NC 文字清洗、標準註解檔頭注入、臥式機型 12 行專屬檔頭置換與安全落盤
"""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

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

    def test_multi_tool_first_tool_dry_cutting_preserves_all_toolpaths(self):
        """
        驗證嚴重安全修正 (項目 2)：
        當多刀 NC 程式的第一把刀為乾切削 (無 M08)、第二把刀才開啟 M08 時，
        extract_body_from_original_nc 與 transform_nc_content_by_profile 絕不可跨刀誤刪第一把刀之刀軌！
        """
        raw_nc_dry_first = """%
O3001
G91 G28 Z0.
T01 M06
S4000 M03
G90 G54 G00 X25.5 Y-30.2
G43 H01 Z50.
G00 Z3.
G01 Z-5.0 F600.
X80.0 Y-30.2
G00 Z50.
M05
G91 G28 Z0.
T02 M06
S8000 M03
G90 G54 G00 X10.0 Y10.0
G43 H02 Z50.
M08
G01 Z-1.5 F1200.
X40.0 Y40.0
M09
M30
%"""
        res_nc = transform_nc_content_by_profile(
            original_nc=raw_nc_dry_first,
            profile_id="FANUC_HORIZONTAL",
            header_info={"drawing_name": "DRY_CUT_TEST"},
            mcs_origin_str="加工原點【G54】：X=MID, Y=MID, Z=TOP",
            raw_ops=[{"tool_number": "T01", "note": "S:4000 F:600", "clean_op_name": "DRY_ROUGH"}],
            program_name="OP10_DRY"
        )
        # 第一把刀的起刀座標與所有切削路徑必須 100% 完整保留
        self.assertIn("G90 G54 G00 X25.5 Y-30.2", res_nc)
        self.assertIn("G00 Z3.", res_nc)
        self.assertIn("G01 Z-5.0 F600.", res_nc)
        self.assertIn("X80.0 Y-30.2", res_nc)
        # 第二把刀的換刀與切削路徑也必須 100% 完整保留
        self.assertIn("T02 M06", res_nc)
        self.assertIn("G01 Z-1.5 F1200.", res_nc)
        self.assertIn("X40.0 Y40.0", res_nc)

    def test_safe_z_ignores_homing_g28_z0(self):
        """
        驗證嚴重安全修正 (項目 3)：
        當原始 NC 無 M06 且開頭含 G91 G28 Z0.、G43 未同行寫 Z 時，
        parse_nc_initial_params 絕不可誤抓 G28 Z0. 作為安全高度 safe_z！
        """
        raw_nc_no_m06 = """%
O3002
G91 G28 Z0.
G90 G54 G00 X12.0 Y34.0
G43 H05
G01 Z-3.0 F500.
M30
%"""
        wcs, tool, speed, safe_z, start_x, start_y = parse_nc_initial_params(
            raw_nc_no_m06, default_wcs="G54", default_tool="5", default_speed="3500", default_safe_z="50."
        )
        self.assertEqual(wcs, "G54")
        self.assertEqual(start_x, "12.0")
        self.assertEqual(start_y, "34.0")
        # 絕不可為 "0."，必須維持安全預設高度 "50."
        self.assertEqual(safe_z, "50.")

    def test_custom_postprocessor_and_normalized_tool_list(self):
        """
        驗證高優先級修正 (項目 4 & 項目 6)：
        1. 選用 "[自訂] MyExternalPost" 與 custom_post_path 時，能自動還原後處理器名稱完成後處理。
        2. 資料夾名稱含 -T01 後綴 (OP10-T01) 而 nc_tasks 檔名為清洗後之 OP10 時，
           process_and_export_nc_tasks 仍能精確匹配工步資料並於 NC 檔頭輸出完整 TOOL LIST。
        """
        class NonJsonCppOp:
            def __init__(self, name, tag):
                self.Name = name
                self.Tag = tag

        cpp_op = NonJsonCppOp("OP10_ROUGH-T01", 9001)
        processed_chunks = [{
            "group_name": "OP10-T01",
            "parent_program": "OP10-T01",
            "stage": "正面",
            "operations": [{
                "op_name": "OP10",
                "clean_op_name": "OP10_ROUGH",
                "program_name": "OP10",
                "tool_number": "T05",
                "tool_spec_display": "E12R0.5",
                "note": "S:4500 F:900",
                "_nx_op_obj": cpp_op,
                "_merged_nx_ops": [cpp_op]
            }]
        }]
        nc_tasks = [{
            "program_name": "OP10",
            "operations": [cpp_op],
            "op_names": ["OP10_ROUGH-T01"],
            "stage": "正面"
        }]

        called_post_names = []
        def mock_post(cam_objects, post_name, output_file, units):
            called_post_names.append(post_name)
            # 模擬 NX 不認識 "[自訂] ..."，只認識清洗後的 "MyExternalPost"
            if post_name.startswith("[自訂]"):
                raise RuntimeError("Unknown postprocessor label")
            with open(output_file, "w", encoding="utf-8") as f:
                f.write("%\nO1010\nT05 M06\nS4500 M03\nG54 G00 X10. Y20.\nG43 H05 Z50.\nM08\nG01 Z-2. F900.\nM30\n%\n")

        mock_cam_setup = MagicMock()
        mock_cam_setup.Postprocess.side_effect = mock_post

        tmp_out = tempfile.mkdtemp()
        succ, fail = process_and_export_nc_tasks(
            cam_setup=mock_cam_setup,
            processed_chunks=processed_chunks,
            profile_id="FANUC_HORIZONTAL",
            header_info={"drawing_name": "PART_CUSTOM_POST"},
            stage_origins={"正面": "加工原點【G54】：X=MID, Y=MID, Z=TOP"},
            output_dir=tmp_out,
            default_post="[自訂] MyExternalPost",
            extension=".nc",
            nc_tasks=nc_tasks,
            custom_post_path="C:/custom_posts/MyExternalPost.pui"
        )
        self.assertEqual(succ, 1)
        self.assertEqual(fail, 0)
        self.assertIn("MyExternalPost", called_post_names)

        out_nc = os.path.join(tmp_out, "OP10.nc")
        self.assertTrue(os.path.exists(out_nc))
        with open(out_nc, "r", encoding="utf-8") as f:
            nc_text = f.read()
        # 驗證 TOOL LIST 完整包含 T05 與 E12R0.5
        self.assertIn("TOOL LIST", nc_text)
        self.assertIn("T05", nc_text)
        self.assertIn("E12R0.5", nc_text)


class TestNcPostOnlyWorkflow(unittest.TestCase):
    """驗證【⚡ 只轉 NC 碼 (不轉工單)】對話框與流程相關修正"""

    def _create_dialog(self, extra_cfg=None):
        import tkinter as tk
        from nc_post_dialog import NcPostDialogApp

        cfg = {
            "nc_tasks": [
                {"program_name": "OP10", "op_count": 2, "op_names": ["ROUGH", "FINISH"]}
            ],
            "default_post": "Fanuc_2026",
            "default_machine_profile": "FANUC_HORIZONTAL",
            "output_dir": tempfile.gettempdir(),
        }
        if extra_cfg:
            cfg.update(extra_cfg)

        root = tk.Tk()
        root.withdraw()
        with patch("machine_profile_manager.get_post_settings", return_value={}):
            app = NcPostDialogApp(root, cfg)
        return root, app

    def test_nc_post_dialog_min_width_fits_four_buttons(self):
        """驗證四顆動作按鈕不會超出視窗最小寬度而被右側裁切"""
        root, app = self._create_dialog()
        try:
            root.update_idletasks()
            min_w, _ = root.minsize()
            # 找出底部按鈕列 (4 個子元件皆為 Button 的 Frame) 的實際需求寬度
            btn_bars = [
                w for w in root.winfo_children()
                if w.winfo_class() == "Frame"
                and len(w.winfo_children()) == 4
                and all(c.winfo_class() == "Button" for c in w.winfo_children())
            ]
            self.assertEqual(len(btn_bars), 1)
            self.assertLessEqual(btn_bars[0].winfo_reqwidth(), min_w)
        finally:
            try:
                root.destroy()
            except Exception:
                pass

    def test_nc_post_dialog_extension_updates_task_listbox_preview(self):
        """驗證切換或輸入副檔名時，待轉出清單預覽檔名即時同步更新"""
        root, app = self._create_dialog()
        try:
            self.assertIn("OP10.nc", app.task_listbox.get(0))
            app.ext_var.set("mpf")
            root.update_idletasks()
            self.assertEqual(app._normalize_extension(), ".mpf")
            self.assertIn("OP10.mpf", app.task_listbox.get(0))
        finally:
            try:
                root.destroy()
            except Exception:
                pass

    def test_nc_post_dialog_empty_output_dir_fallback_and_validation(self):
        """驗證空輸出目錄具備預設退回保護，且手動清空時 on_post_only 會攔截提示"""
        root, app = self._create_dialog({"output_dir": ""})
        try:
            self.assertTrue(bool(app.dir_var.get().strip()))
            app.dir_var.set("   ")
            with patch("nc_post_dialog.messagebox.showwarning") as mock_warn, \
                 patch("machine_profile_manager.save_post_settings"):
                app.on_post_only()
            mock_warn.assert_called_once()
            self.assertEqual(app.result["action"], "cancel")
            self.assertTrue(root.winfo_exists())
        finally:
            try:
                root.destroy()
            except Exception:
                pass

    def test_nc_post_dialog_clears_stale_custom_post_path_on_standard_select(self):
        """驗證從自訂後處理器切回標準後處理器並按只轉 NC 時，會清除舊 custom_post_path"""
        root, app = self._create_dialog({
            "default_post": "[自訂] OldCustom",
            "custom_post_path": "C:/custom/OldCustom.tcl"
        })
        try:
            app.post_var.set("Fanuc_2026")
            with patch("machine_profile_manager.save_post_settings") as mock_save:
                app.on_post_only()
            self.assertEqual(app.result["action"], "post_only")
            self.assertEqual(app.result["postprocessor_name"], "Fanuc_2026")
            self.assertEqual(app.result["custom_post_path"], "")
            mock_save.assert_called_once()
            self.assertEqual(mock_save.call_args.kwargs.get("custom_post_path"), "")
        finally:
            try:
                root.destroy()
            except Exception:
                pass

    def test_nc_post_dialog_cli_auto_action_post_only(self):
        """驗證子行程 CLI 模式 --auto-action post_only 可正確輸出 JSON 結果"""
        with tempfile.TemporaryDirectory() as tmp_dir:
            cfg_path = os.path.join(tmp_dir, "cfg.json")
            res_path = os.path.join(tmp_dir, "res.json")
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump({
                    "nc_tasks": [{"program_name": "OP10", "op_count": 1, "op_names": ["OP_A"]}],
                    "default_post": "Fanuc_2026",
                    "default_machine_profile": "FANUC_HORIZONTAL",
                    "output_dir": tmp_dir
                }, f, ensure_ascii=False)

            proc = subprocess.run(
                [
                    sys.executable,
                    os.path.join(os.path.dirname(__file__), "nc_post_dialog.py"),
                    "--cfg-file", cfg_path,
                    "--res-file", res_path,
                    "--auto-action", "post_only"
                ],
                capture_output=True,
                text=True,
                timeout=20
            )
            self.assertEqual(proc.returncode, 0, msg=proc.stderr)
            with open(res_path, "r", encoding="utf-8") as f:
                res = json.load(f)
            self.assertEqual(res["action"], "post_only")
            self.assertEqual(res["output_dir"], tmp_dir)
            self.assertEqual(res["extension"], ".nc")

    def test_release_nx_collections_clears_containers(self):
        """驗證 _release_nx_collections 可安全清空傳入的集合容器並忽略 None"""
        from nx_cam_to_excel import _release_nx_collections

        lst = [1, 2, 3]
        dct = {"a": 1}
        _release_nx_collections(lst, dct, None)
        self.assertEqual(lst, [])
        self.assertEqual(dct, {})

    def test_nc_post_dialog_browse_output_dir_toggles_topmost_and_normalizes_path(self):
        """驗證變更目錄時會暫時解除 -topmost、指定 parent，並正規化斜線路徑"""
        with tempfile.TemporaryDirectory() as tmp_dir:
            root, app = self._create_dialog({"output_dir": os.path.join(tmp_dir, "nonexistent_sub")})
            try:
                observed_state = {}
                target_forward = os.path.join(tmp_dir, "selected_nc").replace("\\", "/")

                def fake_askdir(**kwargs):
                    observed_state["kwargs"] = kwargs
                    observed_state["topmost_during"] = str(root.attributes("-topmost")).lower() in ("1", "true")
                    return target_forward

                with patch("nc_post_dialog.filedialog.askdirectory", side_effect=fake_askdir):
                    app.browse_output_dir()

                self.assertEqual(observed_state["kwargs"].get("parent"), root)
                self.assertTrue(os.path.isdir(observed_state["kwargs"].get("initialdir")))
                self.assertFalse(observed_state["topmost_during"])
                self.assertTrue(str(root.attributes("-topmost")).lower() in ("1", "true"))
                self.assertEqual(app.dir_var.get(), os.path.normpath(target_forward))
            finally:
                try:
                    root.destroy()
                except Exception:
                    pass

    def test_process_and_export_nc_tasks_normalizes_forward_slash_output_dir(self):
        """驗證傳入含正斜線或引號的輸出目錄時，呼叫 NX Postprocess 使用完全正規化的系統路徑"""
        class OpStub:
            def __init__(self, name, tag):
                self.Name = name
                self.Tag = tag

        op = OpStub("OP10_ROUGH", 101)
        recorded_files = []

        def mock_post(cam_objects, post_name, output_file, units):
            recorded_files.append(output_file)
            with open(output_file, "w", encoding="utf-8") as f:
                f.write("%\nO1010\nT01 M06\nS3000 M03\nG54 G00 X0. Y0.\nG43 H01 Z50.\nM30\n%\n")

        mock_cam_setup = MagicMock()
        mock_cam_setup.Postprocess.side_effect = mock_post

        with tempfile.TemporaryDirectory() as tmp_dir:
            dirty_dir = f' "{tmp_dir.replace(chr(92), "/")}" '
            succ, fail = process_and_export_nc_tasks(
                cam_setup=mock_cam_setup,
                processed_chunks=[{
                    "group_name": "OP10",
                    "parent_program": "OP10",
                    "stage": "正面",
                    "operations": [{
                        "op_name": "OP10",
                        "clean_op_name": "OP10_ROUGH",
                        "program_name": "OP10",
                        "tool_number": "T01",
                        "tool_spec_display": "D10",
                        "_nx_op_obj": op,
                        "_merged_nx_ops": [op]
                    }]
                }],
                profile_id="RAW",
                output_dir=dirty_dir,
                default_post="Fanuc_2026",
                extension=".nc",
                nc_tasks=[{"program_name": "OP10", "operations": [op], "op_names": ["OP10_ROUGH"], "stage": "正面"}]
            )
            self.assertEqual(succ, 1)
            self.assertEqual(fail, 0)
            self.assertEqual(len(recorded_files), 1)
            self.assertEqual(recorded_files[0], os.path.normpath(recorded_files[0]))
            self.assertTrue(os.path.exists(os.path.join(tmp_dir, "OP10.nc")))

    def test_process_and_export_nc_tasks_merges_same_stage_and_disambiguates_cross_stage(self):
        """驗證 H4：同製程同名群組自動合併工序不覆蓋，跨製程同名群組自動附加製程後綴獨立輸出"""
        from nc_processor import process_and_export_nc_tasks

        class OpStub:
            def __init__(self, name, tag):
                self.Name = name
                self.Tag = tag

        op1 = OpStub("OP10_ROUGH_1", 201)
        op2 = OpStub("OP10_ROUGH_2", 202)
        op3 = OpStub("OP10_BACK_1", 203)
        post_calls = []

        def mock_post(cam_objects, post_name, output_file, units):
            post_calls.append(( [o.Name for o in cam_objects], output_file ))
            with open(output_file, "w", encoding="utf-8") as f:
                f.write("%\nO1010\nT01 M06\nS3000 M03\nG54 G00 X0. Y0.\nG43 H01 Z50.\nM30\n%\n")

        mock_cam_setup = MagicMock()
        mock_cam_setup.Postprocess.side_effect = mock_post

        with tempfile.TemporaryDirectory() as tmp_dir:
            succ, fail = process_and_export_nc_tasks(
                cam_setup=mock_cam_setup,
                processed_chunks=[],
                profile_id="RAW",
                output_dir=tmp_dir,
                default_post="Fanuc_2026",
                extension=".nc",
                nc_tasks=[
                    {"program_name": "OP10", "operations": [op1], "op_names": ["OP10_ROUGH_1"], "stage": "正面"},
                    {"program_name": "OP10", "operations": [op2], "op_names": ["OP10_ROUGH_2"], "stage": "正面"},
                    {"program_name": "OP10", "operations": [op3], "op_names": ["OP10_BACK_1"], "stage": "背面"},
                ]
            )
            self.assertEqual(succ, 2)
            self.assertEqual(fail, 0)
            self.assertEqual(len(post_calls), 2)
            # 同製程的兩個同名 OP10 應合併為單一任務，包含 op1 與 op2
            self.assertEqual(post_calls[0][0], ["OP10_ROUGH_1", "OP10_ROUGH_2"])
            self.assertEqual(post_calls[1][0], ["OP10_BACK_1"])
            self.assertTrue(os.path.exists(os.path.join(tmp_dir, "OP10.nc")))
            self.assertTrue(os.path.exists(os.path.join(tmp_dir, "OP10_背面.nc")))

    def test_merge_consecutive_tools_preserves_same_t_num_with_different_tool_spec(self):
        """驗證 H5：相同刀號但不同刀具名稱或刀徑的相鄰工序不得被合併，確保刀號衝突檢測能攔截"""
        import nx_cam_to_excel
        from tool_conflict_detector import detect_tool_conflicts_for_operations

        ops = [
            {
                "op_name": "OP10_MILL",
                "clean_op_name": "OP10_MILL",
                "tool_number": "T01",
                "tool_name": "EM_D10",
                "raw_tool_name": "EM_D10",
                "tool_diameter": 10.0,
                "flute_length": 25.0,
                "holder_length": 50.0,
                "time_seconds": 30.0,
                "stage": "正面",
            },
            {
                "op_name": "OP20_DRILL",
                "clean_op_name": "OP20_DRILL",
                "tool_number": "T01",
                "tool_name": "DRILL_D6",
                "raw_tool_name": "DRILL_D6",
                "tool_diameter": 6.0,
                "flute_length": 30.0,
                "holder_length": 60.0,
                "time_seconds": 20.0,
                "stage": "正面",
            }
        ]
        merged = nx_cam_to_excel.merge_consecutive_tools(ops)
        self.assertEqual(len(merged), 2)
        conflicts = detect_tool_conflicts_for_operations(merged)
        self.assertEqual(len(conflicts), 1)
        self.assertTrue(merged[0].get("has_tool_conflict"))
        self.assertTrue(merged[1].get("has_tool_conflict"))

    def test_evaluate_operations_tool_safety_checks_all_merged_nx_ops(self):
        """驗證 M2：連續同刀合併後，刀具安全評估會遍歷 _merged_nx_ops 中的所有實體工序"""
        import tool_safety_evaluator

        op_first = MagicMock()
        op_first.Name = "OP_SAFE"
        op_second = MagicMock()
        op_second.Name = "OP_COLLISION"

        class FakeChecker:
            def __init__(self):
                self.OperationToCheck = None
                self.CheckNeck = False
                self.NumCutDepthViolations = 0
                self.MaxCutDepthViolation = 0.0

            def PerformCheck(self):
                if self.OperationToCheck is op_second:
                    self.NumCutDepthViolations = 1
                    self.MaxCutDepthViolation = 2.5
                else:
                    self.NumCutDepthViolations = 0
                    self.MaxCutDepthViolation = 0.0

            def Destroy(self):
                pass

        mock_cam_setup = MagicMock()
        mock_cam_setup.CreateCutDepthChecker.side_effect = FakeChecker

        ops_list = [{
            "op_name": "OP_SAFE",
            "tool_number": "T01",
            "flute_length": "20",
            "holder_length": "30",
            "_nx_op_obj": op_first,
            "_merged_nx_ops": [op_first, op_second],
        }]

        warnings = tool_safety_evaluator.evaluate_operations_tool_safety(ops_list, cam_setup=mock_cam_setup)
        self.assertEqual(len(warnings), 1)
        self.assertIn("2.5", ops_list[0].get("tool_safety_warning", ""))

    def test_read_raw_nc_file_supports_utf8_and_cp950(self):
        """驗證 M5：_read_raw_nc_file 同時支援 UTF-8 與 CP950 編碼的原始 NC 暫存檔"""
        from nc_processor import _read_raw_nc_file

        with tempfile.TemporaryDirectory() as tmp_dir:
            cp950_file = os.path.join(tmp_dir, "raw_cp950.nc")
            with open(cp950_file, "wb") as f:
                f.write("%\n(傳統中文註解: 粗銑外型)\nG00 X0. Y0.\nM30\n%\n".encode("cp950"))

            content = _read_raw_nc_file(cp950_file)
            self.assertIn("粗銑外型", content)

    def test_cleanup_session_resources_removes_temp_files_and_clears_collections(self):
        """驗證 H2：_cleanup_session_resources 能清理殘留旗標、暫存截圖並釋放集合"""
        import nx_cam_to_excel

        with tempfile.TemporaryDirectory() as tmp_dir:
            snap_flag = os.path.join(tmp_dir, "_snap_request_1.flag")
            layer_json = os.path.join(tmp_dir, "_layer_info_1.json")
            stage_png = os.path.join(tmp_dir, "_temp_stage_正面.png")
            for p in (snap_flag, layer_json, stage_png):
                with open(p, "w", encoding="utf-8") as f:
                    f.write("temp")

            stage_images = {"正面": stage_png}
            ops_data = [{"op_name": "OP10", "_nx_op_obj": object()}]
            nx_cam_to_excel._cleanup_session_resources(tmp_dir, stage_images, ops_data)

            self.assertFalse(os.path.exists(snap_flag))
            self.assertFalse(os.path.exists(layer_json))
            self.assertFalse(os.path.exists(stage_png))
            self.assertEqual(len(stage_images), 0)
            self.assertEqual(len(ops_data), 0)

    def test_invoke_nc_post_dialog_nonzero_exit_without_res_file_triggers_fallback(self):
        """驗證 H3：nc_post_dialog 子進程異常崩潰且未產生結果檔時，觸發原生備援對話框而非靜默當成僅出工單"""
        import nx_cam_to_excel

        mock_proc = MagicMock()
        mock_proc.returncode = 1
        mock_proc.stderr = "Tkinter initialization failed"

        with patch("subprocess.run", return_value=mock_proc), \
             patch.object(nx_cam_to_excel, "_prompt_nx_native_post_fallback", return_value={"action": "cancel"}) as mock_fb:
            res = nx_cam_to_excel.invoke_nc_post_dialog(
                [{"program_name": "OP10", "op_count": 1, "op_names": ["OP10_ROUGH"], "source_type": "PROGRAM"}],
                "C:/dummy_dir"
            )
            self.assertEqual(res.get("action"), "cancel")
            mock_fb.assert_called_once()


if __name__ == "__main__":
    unittest.main()
