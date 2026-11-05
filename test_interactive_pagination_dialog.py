# -*- coding: utf-8 -*-

"""

單元測試：test_interactive_pagination_dialog.py

驗證工單可視化互動分頁排版編輯器之核心邏輯：

1. 資料模型初始化與智慧多數工段判定 (選項 1.A)

2. 手動切換示圖關聯工段持久性

3. 程式檔名顯示與解析 (首行 📁 檔名，後續 ↳ 同上，無空白列)

4. 檔名群組化切分 (get_page_groups)

5. 點擊刀具之群組聯動選取 (on_tree_select_expand_group)

6. 程式檔名群組整體順序調整 (on_move_group_order)

7. 程式檔名群組跨頁移動與新分頁建立 (move_selected_groups_to_page, on_isolate_group_to_new_page)

8. 單元格直接在序號欄中輸入文字標記，不加流水號 (用戶最新需求 1)

9. 群組截圖欄只對程式群組首行生效，後續刀具為空，且預設為不勾選 (用戶最新需求 2)

10. 單擊截圖欄位直接切換勾選狀態 (用戶最新需求 2)

11. 示圖排版模式：工序後加圖片 vs 圖片統一置於最後一頁

12. 獨立進程與 JSON 原子落盤 (Zero-Debug Protocol)

"""



import unittest

import os

import sys

import json

import tempfile

import subprocess

import tkinter as tk



from interactive_pagination_dialog import (

    InteractivePaginationDialogApp,

    get_item_program_name,

    clean_program_name_helper,

    SEQ_STYLE_PER_PAGE,

    SEQ_STYLE_STAGE_PREFIX,

    SEQ_STYLE_GLOBAL,

    SEQ_STYLE_ZERO_PAD,

    SEQ_STYLE_STEP_10,

    IMG_MODE_AFTER_OPS,

    IMG_MODE_LAST_PAGE

)





class TestInteractivePaginationDialog(unittest.TestCase):



    def setUp(self):

        self.root = tk.Tk()

        self.root.withdraw()



        self.mock_config = {

            "stages": ["M1", "M2"],

            "initial_sequence_style": SEQ_STYLE_PER_PAGE,

            "initial_image_layout": IMG_MODE_AFTER_OPS,

            "default_pages": [

                {

                    "stage": "M1",

                    "operations": [

                        {

                            "program_name": "2205001_ROUGH",

                            "clean_op_name": "OP_M1_1",

                            "stage": "M1",

                            "tool_number": "T01",

                            "time": "00:01:00"

                        },

                        {

                            "program_name": "2205001_ROUGH",

                            "clean_op_name": "OP_M1_2",

                            "stage": "M1",

                            "tool_number": "T02",

                            "time": "00:02:00"

                        },

                        {

                            "program_name": "2205001_FINISH",

                            "clean_op_name": "OP_M1_3",

                            "stage": "M1",

                            "tool_number": "T03",

                            "time": "00:03:00"

                        },

                    ]

                },

                {

                    "stage": "M2",

                    "operations": [

                        {

                            "program_name": "2205002_FINISH",

                            "clean_op_name": "OP_M2_1",

                            "stage": "M2",

                            "tool_number": "T01",

                            "time": "00:01:30"

                        },

                        {

                            "program_name": "2205002_FINISH",

                            "clean_op_name": "OP_M2_2",

                            "stage": "M2",

                            "tool_number": "T04",

                            "time": "00:02:30"

                        },

                    ]

                }

            ]

        }



    def tearDown(self):

        try:

            self.root.destroy()

        except Exception:

            pass



    def test_init_and_majority_stage_detection(self):

        """測試初始載入與多數工段自動偵測 (選項 1.A)"""

        app = InteractivePaginationDialogApp(self.root, self.mock_config)

        self.assertEqual(len(app.pages_data), 2)

        self.assertEqual(app.pages_data[0]["stage"], "M1")

        self.assertEqual(app.pages_data[1]["stage"], "M2")

        self.assertEqual(len(app.pages_data[0]["items"]), 3)

        self.assertEqual(len(app.pages_data[1]["items"]), 2)



    def test_default_capture_is_unchecked(self):

        """

        用戶需求 2 測試：

        截圖勾選預設為不勾選 (need_capture = False)。

        """

        app = InteractivePaginationDialogApp(self.root, self.mock_config)

        # 檢驗所有工步與群組之 need_capture 預設皆為 False

        for page in app.pages_data:

            for op in page["items"]:

                self.assertFalse(op.get("need_capture", False))



        groups = app.get_page_groups(0)

        self.assertFalse(groups[0]["need_capture"])

        self.assertFalse(groups[1]["need_capture"])



    def test_capture_column_only_shows_on_group_first_row(self):

        """

        用戶需求 2 測試：

        截圖勾選欄應該只對「程式群組」生效，而不是全部刀具！

        群組首行顯示截圖狀態 (預設「☐ 略過」)，同群組後續刀具為空字串！

        """

        app = InteractivePaginationDialogApp(self.root, self.mock_config)

        app.current_page_idx = 0

        app.refresh_tree_items()



        # Page 0 包含兩個群組：

        # 群組 1: 2205001_ROUGH (包含 item_0, item_1)

        # 群組 2: 2205001_FINISH (包含 item_2)

        row0 = app.tree.item("item_0")["values"]

        row1 = app.tree.item("item_1")["values"]

        row2 = app.tree.item("item_2")["values"]



        # 截圖欄位位於 index 8 (群組截圖)

        # item_0 是群組首行，顯示 "☐ 略過"

        self.assertEqual(row0[8], "☐ 略過")

        # item_1 是同群組第二把刀，顯示 "" (空白，不單獨顯示勾選框)

        self.assertEqual(row1[8], "")

        # item_2 是第二群組首行，顯示 "☐ 略過"

        self.assertEqual(row2[8], "☐ 略過")



    def test_direct_seq_in_place_edit(self):

        """

        用戶需求 1 測試：

        直接在序號欄中輸入更直觀！

        在序號儲存格啟動內聯編輯直接輸入自訂文字 (如 OP10)，

        序號完全依照輸入文字顯示，不須再加流水序號！

        """

        app = InteractivePaginationDialogApp(self.root, self.mock_config)

        app.current_page_idx = 0

        app.refresh_tree_items()



        # 模擬在 item_0 的序號儲存格啟動內聯編輯並鍵入 "OP10"

        groups = app.get_page_groups(0)

        target_grp = groups[0]  # 2205001_ROUGH



        # 直接模擬內聯輸入 commit 行為

        for op in target_grp["items"]:

            op["custom_seq"] = "OP10"

            op["stage"] = "OP10"

        app.refresh_active_page_details()



        row0 = app.tree.item("item_0")["values"]

        row1 = app.tree.item("item_1")["values"]

        row2 = app.tree.item("item_2")["values"]



        # 序號儲存格以輸入框樣式【 [ 1 ] 】顯示，直觀引導使用者此處具備輸入功能！

        self.assertEqual(str(row0[0]), "[ OP10 ]")

        self.assertEqual(str(row1[0]), "[ OP10 ]")

        # 未自訂文字的第二個群組維持預設流水號樣式 [ 3 ]

        self.assertEqual(str(row2[0]), "[ 3 ]")



        # 檔名欄位顯示包含 [OP10] 標記

        self.assertIn("[OP10]", row0[1])



    def test_click_toggle_group_capture(self):

        """

        用戶需求 2 測試：

        單擊截圖欄位可快速切換勾選狀態 (☐ 略過 <-> ☑ 截圖)。

        """

        app = InteractivePaginationDialogApp(self.root, self.mock_config)

        app.current_page_idx = 0

        app.refresh_tree_items()



        # 原為 False

        groups = app.get_page_groups(0)

        self.assertFalse(groups[0]["need_capture"])



        # 模擬單擊切換

        target_grp = groups[0]

        new_state = not target_grp["need_capture"]

        for op in target_grp["items"]:

            op["need_capture"] = new_state

        app.refresh_active_page_details()



        row0 = app.tree.item("item_0")["values"]

        self.assertIn("☑ 截圖", row0[8])



        # 再次切換

        new_state = not new_state

        for op in target_grp["items"]:

            op["need_capture"] = new_state

        app.refresh_active_page_details()



        row0_cleared = app.tree.item("item_0")["values"]

        self.assertEqual(row0_cleared[8], "☐ 略過")



    def test_image_layout_mode_default_unchecked_no_image(self):

        """

        用戶需求 2 測試：

        預設不勾選截圖時，所有頁面為純刀具頁面 (show_image = False)。

        """

        app = InteractivePaginationDialogApp(self.root, self.mock_config)

        export_pages = app.build_export_pages()



        # 因預設 need_capture 全為 False，所有頁面 show_image 為 False

        for p in export_pages:

            self.assertFalse(p["show_image"])



    def test_image_layout_mode_last_page_with_checked_group(self):

        """

        用戶需求 2 測試：

        示圖排版模式：【圖片統一置於最後一頁】

        當工程師勾選了 M1 的程式群組時，前面頁面為純刀具頁，最後追加 M1 大圖附圖頁！

        """

        app = InteractivePaginationDialogApp(self.root, self.mock_config)

        app.cbo_img_mode.set("2. 圖片統一置於最後一頁 (大圖附圖頁)")



        # 勾選 Page 0 第一個群組截圖

        for op in app.pages_data[0]["items"][:2]:

            op["need_capture"] = True



        export_pages = app.build_export_pages()



        # 產出應包含：第 1 頁 (刀具頁, show_image=False)、第 2 頁 (刀具頁, show_image=False)

        # + 第 3 頁 (M1 專屬大圖附圖頁, show_image=True)

        self.assertEqual(len(export_pages), 3)

        self.assertFalse(export_pages[0]["show_image"])

        self.assertFalse(export_pages[1]["show_image"])



        self.assertTrue(export_pages[2]["is_appendix_image_page"])

        self.assertTrue(export_pages[2]["show_image"])

        self.assertEqual(export_pages[2]["stage"], "M1")



    def test_multiple_operations_checked_produce_separate_screenshots(self):

        """

        用戶核心規則測試：

        「工序右方截圖若有勾選 就是表示該工序需要一張截圖」

        驗證當同一個工段 (如 M1) 內有多個程式群組 (PROG_01 與 PROG_02) 皆被勾選截圖時：

        系統必須為每一個勾選的工序各生成一張專屬截圖任務，不按工段合併！

        """

        app = InteractivePaginationDialogApp(self.root, self.mock_config)



        # 1. 測試模式 B (圖片統一置於最後一頁)

        app.cbo_img_mode.set("2. 圖片統一置於最後一頁 (大圖附圖頁)")

        # 勾選 PROG_01 (前2刀) 與 PROG_02 (後2刀)

        for op in app.pages_data[0]["items"][:2]:

            op["need_capture"] = True

        for op in app.pages_data[0]["items"][2:4]:

            op["need_capture"] = True



        export_pages_b = app.build_export_pages()



        # 檢驗最後生成的附圖頁：應為 2 頁專屬大圖附圖頁

        appendix_pages_b = [p for p in export_pages_b if p.get("is_appendix_image_page")]

        self.assertEqual(len(appendix_pages_b), 1, "勾選2個工序截圖，在模式B下應排於同一附圖頁 (最多6張)！")

        self.assertEqual(appendix_pages_b[0]["target_ops"], ["2205001_ROUGH", "2205001_FINISH"])

        self.assertEqual(appendix_pages_b[0]["image_layout"], "side_by_side")




        # 2. 測試模式 A (工序後加圖片)

        # 用戶最新核心規則：「雙工序左右並排同頁」與「工序刀具 緊接著工序截圖」

        # (1) 當同一頁勾選 2 個工序且刀數 <= 20 刀：自動啟用雙工序左右並排 (side_by_side)，同頁容納！

        app.cbo_img_mode.set("1. 工序後加圖片 (同頁下方/工序緊隨)")

        export_pages_a = app.build_export_pages()

        image_pages_a = [p for p in export_pages_a if p.get("show_image")]

        self.assertEqual(len(image_pages_a), 1, "模式A同頁2工序勾選截圖時，應生成1頁雙圖左右並排頁！")

        self.assertEqual(image_pages_a[0]["image_layout"], "side_by_side")

        self.assertEqual(image_pages_a[0]["target_ops"], ["2205001_ROUGH", "2205001_FINISH"])

        self.assertEqual(len(image_pages_a[0]["rows"]), 3, "包含兩道工序之所有刀具 (2刀 + 1刀)")

        self.assertFalse(image_pages_a[0]["is_appendix_image_page"])



        # (2) 當使用者將第二工序移至下一頁：各自生成單圖全幅緊隨頁面

        app.pages_data[0]["items"] = app.pages_data[0]["items"][:2]  # 僅留 ROUGH

        app.pages_data.append({"stage": "M1", "items": [

            {"program_name": "2205001_FINISH", "clean_op_name": "OP_M1_3", "stage": "M1", "tool_number": "T03", "time": "00:03:00", "need_capture": True}

        ]})

        export_pages_split = app.build_export_pages()

        split_image_pages = [p for p in export_pages_split if p.get("show_image")]

        self.assertEqual(len(split_image_pages), 2, "跨頁時各自生成單圖獨立頁！")

        self.assertEqual(split_image_pages[0]["image_layout"], "single")

        self.assertEqual(split_image_pages[0]["target_op"], "2205001_ROUGH")

        self.assertEqual(split_image_pages[1]["image_layout"], "single")

        self.assertEqual(split_image_pages[1]["target_op"], "2205001_FINISH")



    def test_move_group_order(self):

        """測試以程式檔名為群組整體上下調整順序"""

        app = InteractivePaginationDialogApp(self.root, self.mock_config)

        app.current_page_idx = 0

        app.refresh_tree_items()



        app.tree.selection_set("item_0", "item_1")

        app.on_move_group_order(1)



        new_groups = app.get_page_groups(0)

        self.assertEqual(new_groups[0]["prog_name"], "2205001_FINISH")

        self.assertEqual(new_groups[1]["prog_name"], "2205001_ROUGH")



    def test_move_group_cross_page(self):

        """測試選取群組完整跨頁移至下一頁"""

        app = InteractivePaginationDialogApp(self.root, self.mock_config)

        app.current_page_idx = 0

        app.refresh_tree_items()



        app.tree.selection_set("item_0", "item_1")

        app.on_move_group_to_next_page()



        self.assertEqual(len(app.pages_data[0]["items"]), 1)

        self.assertEqual(len(app.pages_data[1]["items"]), 4)



    def test_move_group_to_next_page_at_last_page_creates_new_page_and_moves_group(self):
        """測試在最後一頁點擊移至下一頁時，確認後自動建立新分頁並將選取群組完整移入"""
        from unittest.mock import patch
        app = InteractivePaginationDialogApp(self.root, self.mock_config)
        # 移除第 2 頁，使第 1 頁 (index 0) 成為最後一頁 (包含 ROUGH 2刀 與 FINISH 1刀)
        app.pages_data.pop(1)
        app.refresh_all()

        orig_page_count = len(app.pages_data)
        self.assertEqual(orig_page_count, 1)
        self.assertEqual(len(app.pages_data[0]['items']), 3)

        # 選取第二個程式群組 (item_2: 2205001_FINISH)
        app.tree.selection_set('item_2')

        with patch('tkinter.messagebox.askyesno', return_value=True):
            app.on_move_group_to_next_page()

        # 驗證總頁數增加 1 頁
        self.assertEqual(len(app.pages_data), 2)
        # 驗證原頁面扣除移出的工序 (剩餘 ROUGH 2刀)
        self.assertEqual(len(app.pages_data[0]['items']), 2)
        # 驗證新分頁完整移入 FINISH 群組 (1刀)
        self.assertEqual(len(app.pages_data[1]['items']), 1)
        self.assertEqual(app.pages_data[1]['items'][0]['program_name'], '2205001_FINISH')
        # 驗證當前頁面焦點切換至新分頁
        self.assertEqual(app.current_page_idx, 1)

    def test_subprocess_auto_confirm_atomic_output(self):

        """測試透過子進程獨立啟動並原子輸出 JSON (Zero-Debug Protocol)"""

        with tempfile.TemporaryDirectory() as tmp_dir:

            cfg_path = os.path.join(tmp_dir, "test_cfg.json")

            res_path = os.path.join(tmp_dir, "test_res.json")



            with open(cfg_path, "w", encoding="utf-8") as f:

                json.dump(self.mock_config, f, ensure_ascii=False)



            script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "interactive_pagination_dialog.py")



            cmd = [

                sys.executable,

                script_path,

                "--cfg-file", cfg_path,

                "--res-file", res_path,

                "--auto-action", "ok"

            ]



            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=10)

            self.assertEqual(proc.returncode, 0)

            self.assertTrue(os.path.exists(res_path))



            with open(res_path, "r", encoding="utf-8") as f:

                data = json.load(f)



            self.assertEqual(data.get("action"), "ok")

            self.assertEqual(data.get("sequence_style"), SEQ_STYLE_PER_PAGE)





    def test_option_c_greedy_packing_within_capacity(self):
        """
        用戶核心規則測試 (選項 C)：
        「盡可能於排版可接受的範圍內(工序+截圖) 排成同一頁 排不下的再建立分頁」
        驗證：
        1. 3個工序群組在同頁，其中2個勾選截圖、1個未勾選 (共 12 刀)：
           - 12 刀 <= 20 刀，且 2 張截圖 <= 2 張上限
           - 必須 100% 排在同一頁 (side_by_side)，絕不可擅自拆成 3 頁！
        2. 3個工序群組皆勾選截圖 (共 12 刀)：
           - 單頁最多容納 2 張圖片，因此第 1 頁容納前 2 工序與雙圖
           - 第 3 個工序排不下，自動建立第 2 頁容納該工序與其單圖 (總共 2 頁，而非 3 頁)！
        3. 2個工序群組皆勾選截圖，但總刀數為 24 刀 (> 20 刀)：
           - 超過同頁刀圖容納上限 (20 刀)，自動分為 2 頁各自展示！
        """
        app = InteractivePaginationDialogApp(self.root, self.mock_config)
        app.cbo_img_mode.set("1. 工序後加圖片 (同頁下方/工序緊隨)")

        # 情境 1: 3 群組 (G1=4刀[圖], G2=4刀[圖], G3=4刀[無圖]) -> 12 刀 + 2 圖 -> 應排為 1 頁
        app.pages_data = [{
            "page_id": 1,
            "stage": "M1",
            "items": [
                {"program_name": "OP10", "clean_op_name": "T1", "tool_number": "T01", "need_capture": True},
                {"program_name": "OP10", "clean_op_name": "T2", "tool_number": "T02", "need_capture": True},
                {"program_name": "OP10", "clean_op_name": "T3", "tool_number": "T03", "need_capture": True},
                {"program_name": "OP10", "clean_op_name": "T4", "tool_number": "T04", "need_capture": True},
                {"program_name": "OP20", "clean_op_name": "T5", "tool_number": "T05", "need_capture": True},
                {"program_name": "OP20", "clean_op_name": "T6", "tool_number": "T06", "need_capture": True},
                {"program_name": "OP20", "clean_op_name": "T7", "tool_number": "T07", "need_capture": True},
                {"program_name": "OP20", "clean_op_name": "T8", "tool_number": "T08", "need_capture": True},
                {"program_name": "OP30", "clean_op_name": "T9", "tool_number": "T09", "need_capture": False},
                {"program_name": "OP30", "clean_op_name": "T10", "tool_number": "T10", "need_capture": False},
                {"program_name": "OP30", "clean_op_name": "T11", "tool_number": "T11", "need_capture": False},
                {"program_name": "OP30", "clean_op_name": "T12", "tool_number": "T12", "need_capture": False},
            ]
        }]
        pages_1 = app.build_export_pages()
        self.assertEqual(len(pages_1), 1, "12 刀且 2 張截圖，必須完整容納於同一頁！")
        self.assertEqual(len(pages_1[0]["rows"]), 12)
        self.assertTrue(pages_1[0]["show_image"])
        self.assertEqual(pages_1[0]["image_layout"], "side_by_side")
        self.assertEqual(pages_1[0]["target_ops"], ["OP10", "OP20"])

        # 情境 2: 3 群組皆需截圖 (G1=4刀[圖], G2=4刀[圖], G3=4刀[圖]) -> 單頁容納上限2圖，第3圖排不下自動分頁 -> 共 2 頁
        app.pages_data[0]["items"][8]["need_capture"] = True
        pages_2 = app.build_export_pages()
        self.assertEqual(len(pages_2), 2, "3 個需截圖工序，應以 2 頁排完 (P1雙圖, P2單圖)，不可強拆為 3 頁！")
        self.assertEqual(len(pages_2[0]["rows"]), 8)
        self.assertEqual(pages_2[0]["image_layout"], "side_by_side")
        self.assertEqual(pages_2[0]["target_ops"], ["OP10", "OP20"])
        self.assertEqual(len(pages_2[1]["rows"]), 4)
        self.assertEqual(pages_2[1]["image_layout"], "single")
        self.assertEqual(pages_2[1]["target_op"], "OP30")

        # 情境 3: 2 群組皆需截圖，但總刀數 24 刀 (> 20 刀，放不下同頁示圖) -> 自動分頁
        op10_items = [{"program_name": "OP10", "clean_op_name": f"T{i}", "tool_number": f"T{i:02d}", "need_capture": True} for i in range(1, 13)]
        op20_items = [{"program_name": "OP20", "clean_op_name": f"T{i}", "tool_number": f"T{i:02d}", "need_capture": True} for i in range(13, 25)]
        app.pages_data = [{"page_id": 1, "stage": "M1", "items": op10_items + op20_items}]
        pages_3 = app.build_export_pages()
        self.assertEqual(len(pages_3), 2, "總刀數 24 刀放不下同頁示圖，自動分 2 頁各自展示！")
        self.assertEqual(len(pages_3[0]["rows"]), 12)
        self.assertEqual(pages_3[0]["target_op"], "OP10")
        self.assertEqual(len(pages_3[1]["rows"]), 12)
        self.assertEqual(pages_3[1]["target_op"], "OP20")



    def test_mode_b_packing_up_to_6_images_per_page(self):
        """
        用戶最新規則測試：
        【2. 圖片統一置於最後一頁 (大圖附圖頁)】
        「此選項 目前會將圖片至於最後 但是會單獨分頁 應修正成 盡可能將圖片置於同一頁，每頁最多6張示圖」
        驗證：
        1. 勾選 3 個工序截圖 -> 1 頁附圖頁 (grid 佈局)
        2. 勾選 6 個工序截圖 -> 1 頁附圖頁 (grid 佈局，達到單頁上限 6 張)
        3. 勾選 7 個工序截圖 -> 超過單頁 6 張上限，自動分 2 頁附圖頁 (第1頁 6 張 grid, 第2頁 1 張 single)
        """
        app = InteractivePaginationDialogApp(self.root, self.mock_config)
        app.cbo_img_mode.set("2. 圖片統一置於最後一頁 (大圖附圖頁)")

        # 情境 1: 3 個工序勾選截圖
        items_3 = [
            {"program_name": f"OP{i*10}", "clean_op_name": f"T{i}", "tool_number": f"T{i:02d}", "need_capture": True}
            for i in range(1, 4)
        ]
        app.pages_data = [{"page_id": 1, "stage": "M1", "items": items_3}]
        pages_3 = app.build_export_pages()
        app_pages_3 = [p for p in pages_3 if p.get("is_appendix_image_page")]
        self.assertEqual(len(app_pages_3), 1, "3個截圖應收攏於同一附圖頁！")
        self.assertEqual(app_pages_3[0]["image_layout"], "grid")
        self.assertEqual(len(app_pages_3[0]["target_ops"]), 3)

        # 情境 2: 6 個工序勾選截圖
        items_6 = [
            {"program_name": f"OP{i*10}", "clean_op_name": f"T{i}", "tool_number": f"T{i:02d}", "need_capture": True}
            for i in range(1, 7)
        ]
        app.pages_data = [{"page_id": 1, "stage": "M1", "items": items_6}]
        pages_6 = app.build_export_pages()
        app_pages_6 = [p for p in pages_6 if p.get("is_appendix_image_page")]
        self.assertEqual(len(app_pages_6), 1, "6個截圖應剛好收攏於同一附圖頁！")
        self.assertEqual(app_pages_6[0]["image_layout"], "grid")
        self.assertEqual(len(app_pages_6[0]["target_ops"]), 6)

        # 情境 3: 7 個工序勾選截圖 -> 超過上限 6 張，第 7 張應建立新附圖頁 (共 2 頁附圖頁)
        items_7 = [
            {"program_name": f"OP{i*10}", "clean_op_name": f"T{i}", "tool_number": f"T{i:02d}", "need_capture": True}
            for i in range(1, 8)
        ]
        app.pages_data = [{"page_id": 1, "stage": "M1", "items": items_7}]
        pages_7 = app.build_export_pages()
        app_pages_7 = [p for p in pages_7 if p.get("is_appendix_image_page")]
        self.assertEqual(len(app_pages_7), 2, "7個截圖超過單頁6張上限，應分為2頁附圖頁！")
        self.assertEqual(app_pages_7[0]["image_layout"], "grid")
        self.assertEqual(len(app_pages_7[0]["target_ops"]), 6)
        self.assertEqual(app_pages_7[1]["image_layout"], "single")
        self.assertEqual(len(app_pages_7[1]["target_ops"]), 1)
        self.assertEqual(app_pages_7[1]["target_ops"][0], "OP70")

def test_seq_cell_input_and_clear_behavior(self):
        """
        用戶最新需求測試：「序號欄可輸入文字功能 可以讓輸入欄顯示在 序號列嗎? 如此可直觀了解序號欄有輸入功能 並可移除提示列 精簡版面」
        驗證：
        1. 序號儲存格預設呈現【 [ 1 ] 】、【 [ 2 ] 】輸入框外觀。
        2. 表頭標示為「序號 / 標記 [可輸入 ✎]」。
        3. 上方提示列 (group_edit_bar) 已完全移除 (self.ent_custom_seq 為 None)。
        4. 清空輸入內容時，自動清空 custom_seq 並恢復預設流水號樣式。
        """
        app = InteractivePaginationDialogApp(self.root, self.mock_config)
        app.current_page_idx = 0
        app.refresh_tree_items()

        # 斷言 1：表頭更新
        heading_text = app.tree.heading("seq")["text"]
        self.assertIn("可點擊輸入", heading_text)

        # 斷言 2：上方提示列已移除
        self.assertIsNone(app.ent_custom_seq)
        self.assertIsNone(app.lbl_selected_group_name)

        # 斷言 3：初始儲存格為 [ 1 ]
        row0_val = app.tree.item("item_0")["values"][0]
        self.assertEqual(row0_val, "[ 1 ]")

        # 斷言 4：清除文字時恢復預設序號
        groups = app.get_page_groups(0)
        target_grp = groups[0]
        target_grp["custom_seq"] = "TEMP_VAL"
        app.refresh_active_page_details()
        self.assertEqual(app.tree.item("item_0")["values"][0], "[ TEMP_VAL ]")

        # 模擬清空文字
        target_grp["custom_seq"] = ""
        app.refresh_active_page_details()
        self.assertEqual(app.tree.item("item_0")["values"][0], "[ 1 ]")

if __name__ == "__main__":

    unittest.main()

