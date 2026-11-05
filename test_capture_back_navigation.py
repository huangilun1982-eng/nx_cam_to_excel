# -*- coding: utf-8 -*-
import unittest
from unittest.mock import MagicMock, patch
import os
import sys
import tempfile
import subprocess

for mod in ["NXOpen", "NXOpen.CAM", "NXOpen.UF", "NXOpen.Gateway"]:
    if mod not in sys.modules:
        sys.modules[mod] = MagicMock()

from nx_cam_to_excel import run_interactive_capture_wizard


class TestCaptureBackNavigation(unittest.TestCase):
    """
    截圖階段「回到上一步」與「顯示該截圖用於哪一個工序」功能專屬測試：
    1. 驗證非首工序 (s_idx > 0) 點選上一步時，可回退到上一工序重新設定或拍攝。
    2. 驗證首工序 (s_idx == 0) 點選上一步時，向主進程發出 BACK_TO_PAGINATION 導航信號。
    3. 驗證截圖階段精準顯示該截圖用於哪一個工序 (包含刀具規格與工單版面位置)。
    """

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.mock_uf = MagicMock()
        self.mock_uf.Layer.CycleByLayer.return_value = 0  # 避免 mock 物件在 CycleByLayer while 迴圈陷入死循環
        self.mock_ui = MagicMock()
        self.mock_listing = MagicMock()
        self.mock_part = MagicMock()

        self.dummy_script = os.path.join(self.test_dir, "capture_assistant_gui.py")
        with open(self.dummy_script, "w", encoding="utf-8") as f:
            f.write("# dummy script")

    def tearDown(self):
        import shutil
        if os.path.exists(self.test_dir):
            try:
                shutil.rmtree(self.test_dir)
            except Exception:
                pass

    @patch("subprocess.Popen")
    @patch("nx_cam_to_excel.resolve_asset_file")
    @patch("nx_cam_to_excel.resolve_python_runtime")
    @patch("nx_cam_to_excel.capture_nx_viewport")
    def test_back_button_revisits_previous_operation(
        self, mock_snap, mock_runtime, mock_asset, mock_popen
    ):
        """
        測試工序間回退：
        有兩道工序 ["755613070-PEEK", "755613070-A"]：
        1. PEEK 拍照成功 (returncode=0) -> 進入 A
        2. A 點選「⬅ 上一步」 (returncode=3) -> 回退到 PEEK
        3. PEEK 重新拍照 (returncode=0) -> 再次進入 A
        4. A 拍照成功 (returncode=0) -> 完成全部截圖
        """
        mock_asset.return_value = self.dummy_script
        mock_runtime.return_value = (["pythonw.exe"], {})
        mock_snap.return_value = True

        stages = ["755613070-PEEK", "755613070-A"]

        # 模擬 Popen 的 returncode 依序為：0 (拍PEEK), 3 (A按上一步), 0 (重拍PEEK), 0 (拍A)
        proc_returns = [0, 3, 0, 0]
        proc_mocks = []
        for code in proc_returns:
            p = MagicMock()
            p.poll.return_value = code
            p.returncode = code
            proc_mocks.append(p)

        mock_popen.side_effect = proc_mocks

        res_images = run_interactive_capture_wizard(
            stages, self.test_dir, self.mock_uf, self.mock_ui, self.mock_listing, self.mock_part
        )

        self.assertEqual(mock_popen.call_count, 4)
        self.assertIn("755613070-PEEK", res_images)
        self.assertIn("755613070-A", res_images)

    @patch("subprocess.Popen")
    @patch("nx_cam_to_excel.resolve_asset_file")
    @patch("nx_cam_to_excel.resolve_python_runtime")
    def test_back_button_at_first_operation_returns_to_pagination(
        self, mock_runtime, mock_asset, mock_popen
    ):
        """
        測試首工序回退：
        在第一個工序點選「⬅ 上一步」時，必須回傳 BACK_TO_PAGINATION 導航字典。
        """
        mock_asset.return_value = self.dummy_script
        mock_runtime.return_value = (["pythonw.exe"], {})

        mock_proc = MagicMock()
        mock_proc.poll.return_value = 3
        mock_proc.returncode = 3 # 點選上一步
        mock_popen.return_value = mock_proc

        stages = ["755613070-PEEK"]

        result = run_interactive_capture_wizard(
            stages, self.test_dir, self.mock_uf, self.mock_ui, self.mock_listing, self.mock_part
        )

        self.assertIsInstance(result, dict)
        self.assertEqual(result.get("__ACTION__"), "BACK_TO_PAGINATION")

    def test_capture_assistant_gui_structure_and_back_logic(self):
        """
        驗證 capture_assistant_gui.py 中確實包含「⬅ 上一步」/「⬅ 返回排版」按鈕邏輯與 returncode 3
        """
        gui_path = r"c:\NX_Standard\Template\capture_assistant_gui.py"
        self.assertTrue(os.path.exists(gui_path))

        with open(gui_path, "r", encoding="utf-8") as f:
            code = f.read()

        self.assertIn("def on_back():", code)
        self.assertIn("exit_code[0] = 3", code)
        self.assertIn('" ⬅ 返回排版 "', code)
        self.assertIn('" ⬅ 上一步 "', code)

        idx_back = code.find("btn_back.pack")
        idx_snap = code.find("btn_confirm.pack")
        self.assertTrue(idx_back < idx_snap, "上一步按鈕必須位於立即拍照左側！")

    @patch("subprocess.Popen")
    @patch("nx_cam_to_excel.resolve_asset_file")
    @patch("nx_cam_to_excel.resolve_python_runtime")
    @patch("nx_cam_to_excel.capture_nx_viewport")
    def test_capture_assistant_displays_target_operation_and_details(
        self, mock_snap, mock_runtime, mock_asset, mock_popen
    ):
        """
        用戶最新需求測試：「於截圖階段 顯示該截圖用於哪一個工序」
        驗證：
        1. run_interactive_capture_wizard 正確向小工具傳送 --target-op, --tools, --placement
        2. ListingWindow 清楚印出該截圖用於哪一個工序
        3. capture_assistant_gui.py 包含「🎯 此截圖用於工序：【...】」之高對比資訊卡
        """
        mock_asset.return_value = self.dummy_script
        mock_runtime.return_value = (["pythonw.exe"], {})
        mock_snap.return_value = True

        mock_proc = MagicMock()
        mock_proc.poll.return_value = 0
        mock_proc.returncode = 0
        mock_popen.return_value = mock_proc

        structured_stages = [
            {
                "target_op": "755613070-PEEK",
                "stage": "通用工段",
                "custom_seq": "OP10",
                "tools_summary": "T14(E4), T02(E2)",
                "placement": "工單第 1 頁 (刀具下方示圖區)"
            }
        ]

        run_interactive_capture_wizard(
            structured_stages, self.test_dir, self.mock_uf, self.mock_ui, self.mock_listing, self.mock_part
        )

        # 斷言 1：Popen 調用參數中包含專屬 target-op 與詳細工序說明
        called_cmd = mock_popen.call_args[0][0]
        self.assertIn("--target-op", called_cmd)
        self.assertIn("755613070-PEEK", called_cmd)
        self.assertIn("--tools", called_cmd)
        self.assertIn("T14(E4), T02(E2)", called_cmd)
        self.assertIn("--placement", called_cmd)
        self.assertIn("工單第 1 頁 (刀具下方示圖區)", called_cmd)

        # 斷言 2：ListingWindow 日誌指明該截圖用於哪一個工序
        listing_calls = [c[0][0] for c in self.mock_listing.WriteLine.call_args_list]
        self.assertTrue(any("示圖用於工序【755613070-PEEK】" in msg for msg in listing_calls))

        # 斷言 3：capture_assistant_gui.py 中包含醒目的工序用途標籤
        gui_path = r"c:\NX_Standard\Template\capture_assistant_gui.py"
        with open(gui_path, "r", encoding="utf-8") as f:
            gui_code = f.read()
        self.assertIn("🎯 此截圖用於工序：【{target_op}】", gui_code)
        self.assertIn("🔧 包含刀具：", gui_code)



    def test_overlay_op_label_on_image(self):
        """
        用戶最新需求測試：「有辦法讓工序截圖上顯示工序名稱嗎?」
        驗證：
        1. overlay_op_label_on_image 能精準在圖片左上角壓印純工序名稱 (如 XXXX) 之深藍膠囊標記
        2. 當工段為通用工段時，標記為「加工示圖 / 裝夾全貌」
        3. 輸出檔案健全且合法 (非空且可正常讀取)
        """
        from nx_cam_to_excel import overlay_op_label_on_image
        from PIL import Image

        test_img = os.path.join(self.test_dir, "test_crop.png")
        Image.new("RGB", (800, 600), (255, 255, 255)).save(test_img)

        # 測試 1：標準工序名稱
        ok = overlay_op_label_on_image(test_img, "755613070-PEEK", listing=self.mock_listing)
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(test_img))
        self.assertGreater(os.path.getsize(test_img), 100)

        # 檢驗 Listing 記錄
        listing_calls = [c[0][0] for c in self.mock_listing.WriteLine.call_args_list]
        self.assertTrue(any("755613070-PEEK" in msg for msg in listing_calls))

        # 測試 2：通用工段名稱替換
        test_img_gen = os.path.join(self.test_dir, "test_general.png")
        Image.new("RGB", (800, 600), (255, 255, 255)).save(test_img_gen)
        ok_gen = overlay_op_label_on_image(test_img_gen, "通用工段", listing=self.mock_listing)
        self.assertTrue(ok_gen)
        listing_calls_gen = [c[0][0] for c in self.mock_listing.WriteLine.call_args_list]
        self.assertTrue(any("加工示圖 / 裝夾全貌" in msg for msg in listing_calls_gen))


    def test_smart_crop_image(self):
        """
        用戶最新需求測試：「圖片智慧裁切/動態延展以放大畫面特徵」
        驗證：
        1. smart_crop_image 智慧邊界裁切：有效偵測外圍大面積純白背景並進行安全裁切
        2. 裁切後長寬緊湊貼齊，保留 4% 安全呼吸外框，不切除工件實體
        3. 全白背景安全防禦，不損壞圖檔
        """
        from nx_cam_to_excel import smart_crop_image
        from PIL import Image, ImageDraw

        test_img = os.path.join(self.test_dir, "test_wide_viewport.png")
        # 模擬 1920x1080 寬螢幕截圖，工件僅在中間 500x350 區域
        im = Image.new("RGB", (1920, 1080), (255, 255, 255))
        draw = ImageDraw.Draw(im)
        draw.rectangle([700, 360, 1200, 710], fill=(80, 120, 160))
        im.save(test_img)

        with Image.open(test_img) as _tmp_im: orig_w, orig_h = _tmp_im.size
        self.assertEqual((orig_w, orig_h), (1920, 1080))

        ok = smart_crop_image(test_img, listing=self.mock_listing)
        self.assertTrue(ok)

        # 檢驗裁切後尺寸
        with Image.open(test_img) as cropped_im:
            cw, ch = cropped_im.size
            self.assertLess(cw, 800, "寬度應由 1920 顯著緊縮至工件主體大小")
            self.assertLess(ch, 600, "高度應由 1080 顯著緊縮至工件主體大小")
            self.assertGreater(cw, 500, "寬度必須完整保留工件幾何特徵 (500寬)")
            self.assertGreater(ch, 350, "高度必須完整保留工件幾何特徵 (350高)")

    def test_post_process_captured_image_fast_pipeline(self):
        """
        驗證一體化記憶體影像處理管線 post_process_captured_image：
        1. 在單次記憶體走訪中完成 Alpha 純白底合成、智慧邊界裁切與工序標記膠囊壓印。
        2. 原圖為 1920x1080，工件實體為 500x350，裁切後長寬緊縮並保留安全邊距。
        3. 輸出檔案合法且包含深藍工序標籤。
        """
        from nx_cam_to_excel import post_process_captured_image
        from PIL import Image, ImageDraw

        test_img = os.path.join(self.test_dir, "test_pipeline.png")
        # 建立具有透明度通道之 RGBA 圖像，模擬 NX 視圖導出
        im = Image.new("RGBA", (1920, 1080), (255, 255, 255, 0))
        draw = ImageDraw.Draw(im)
        draw.rectangle([700, 360, 1200, 710], fill=(80, 120, 160, 255))
        im.save(test_img, "PNG")

        ok = post_process_captured_image(test_img, op_name="OP10-ROUGH", listing=self.mock_listing)
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(test_img))

        with Image.open(test_img) as res_im:
            self.assertEqual(res_im.mode, "RGB", "輸出圖檔應為 RGB 純白底色")
            rw, rh = res_im.size
            self.assertLess(rw, 800)
            self.assertLess(rh, 600)
            self.assertGreater(rw, 500)
            self.assertGreater(rh, 350)

    def test_ensure_white_background_image_alpha_composite(self):
        """
        驗證輕量化 ensure_white_background_image：
        1. 針對 RGBA 透明圖檔合成至純白底色。
        2. 執行迅速且無死循環或 floodfill 耗時。
        """
        from nx_cam_to_excel import ensure_white_background_image
        from PIL import Image

        test_img = os.path.join(self.test_dir, "test_alpha.png")
        im = Image.new("RGBA", (400, 400), (0, 0, 0, 0)) # 全透明圖檔
        im.save(test_img, "PNG")

        ensure_white_background_image(test_img)
        self.assertTrue(os.path.exists(test_img))

        with Image.open(test_img) as res_im:
            self.assertEqual(res_im.mode, "RGB")
            # 檢查角落是否已變成純白色 (255, 255, 255)
            self.assertEqual(res_im.getpixel((0, 0)), (255, 255, 255))

    def test_export_multipage_no_cross_operation_image_pollution(self):
        """
        驗證修復多頁匯出跨工序圖片污染問題：
        當 stage_images 僅有 1 張特定工序圖片 (例如 OP10)，而當前分頁為 OP20 且無圖片時，
        不可將 OP10 的圖片粗暴回退誤植入 OP20 分頁！
        """
        from nx_cam_to_excel import export_multipage_via_vbs
        import tempfile

        fake_tpl = os.path.join(self.test_dir, "template.xlsx")
        with open(fake_tpl, "w", encoding="utf-8") as f:
            f.write("fake xlsx")
        out_xlsx = os.path.join(self.test_dir, "out.xlsx")

        op10_img = os.path.join(self.test_dir, "op10.png")
        with open(op10_img, "w") as f:
            f.write("fake png")

        stage_images = {"OP10": op10_img}

        pages = [
            {
                "page_num": 1,
                "stage": "OP10",
                "target_op": "OP10",
                "target_ops": ["OP10"],
                "show_image": True,
                "image_layout": "single",
                "image_top_row": 18,
                "image_bottom_row": 36,
                "is_appendix": False,
                "rows": []
            },
            {
                "page_num": 2,
                "stage": "OP20",
                "target_op": "OP20",
                "target_ops": ["OP20"],
                "show_image": True,
                "image_layout": "single",
                "image_top_row": 18,
                "image_bottom_row": 36,
                "is_appendix": False,
                "rows": []
            }
        ]

        mock_work_part = MagicMock()
        mock_work_part.Leaf = "TEST_PART"
        mock_work_part.FullPath = os.path.join(self.test_dir, "TEST_PART.prt")

        captured_vbs = []
        def fake_sub_run(cmd, *args, **kwargs):
            vbs_p = cmd[2]
            if os.path.exists(vbs_p):
                with open(vbs_p, "r", encoding="cp950", errors="ignore") as vf:
                    captured_vbs.append(vf.read())
            return MagicMock(returncode=0)

        with patch("subprocess.run", side_effect=fake_sub_run):
            export_multipage_via_vbs(
                pages=pages,
                template_path=fake_tpl,
                output_path=out_xlsx,
                work_part=mock_work_part,
                header_info={},
                stage_images=stage_images
            )

        self.assertTrue(len(captured_vbs) > 0)
        vbs_code = captured_vbs[0]
        self.assertIn("op10.png", vbs_code)
        p2_idx = vbs_code.find('ws.Name = "第2頁"')
        self.assertGreater(p2_idx, 0)
        p2_section = vbs_code[p2_idx:]
        self.assertNotIn("Pictures.Insert", p2_section, "OP20 未拍照時，第2頁絕不可被 OP10 圖片污染誤植！")



    def test_batch_overlay_op_labels(self):
        """
        用戶最新需求測試：「確認 截圖自動標記名稱 功能是否正常作用?」
        驗證：
        1. batch_overlay_op_labels 能單次批次完成多張示圖的工序名稱標籤壓印。
        2. 正確處理標準工序與通用工段名稱。
        3. 快取機制生效，已壓印之圖片第二次呼叫自動略過，避免重複壓印。
        4. 外部 Python 備援模式能正確將 JSON 任務交由單一進程批次完成。
        """
        from nx_cam_to_excel import batch_overlay_op_labels
        from PIL import Image

        img1 = os.path.join(self.test_dir, "batch_op10.png")
        img2 = os.path.join(self.test_dir, "batch_op20.png")
        img3 = os.path.join(self.test_dir, "batch_general.png")

        Image.new("RGB", (800, 600), (255, 255, 255)).save(img1)
        Image.new("RGB", (1024, 768), (255, 255, 255)).save(img2)
        Image.new("RGB", (640, 480), (255, 255, 255)).save(img3)

        stage_images = {
            "OP10-FACE": img1,
            "OP20-DRILL": img2,
            "通用工段": img3
        }

        # 1. 執行批次壓印
        ok = batch_overlay_op_labels(stage_images, listing=self.mock_listing, force=True)
        self.assertTrue(ok)

        # 檢驗圖檔左上角膠囊像素 (margin=24, pad=16, 位於 (30, 30) 應為深藍色 #004B87)
        for p in [img1, img2, img3]:
            with Image.open(p) as im:
                px = im.getpixel((30, 30))
                self.assertEqual(px, (0, 75, 135), f"示圖 {p} 左上角應壓印深藍色膠囊 (#004B87)")

        # 2. 檢驗防重複壓印快取
        self.mock_listing.reset_mock()
        ok_cached = batch_overlay_op_labels(stage_images, listing=self.mock_listing, force=False)
        self.assertTrue(ok_cached)
        # 由於已快取，Listing 不會再次重複壓印日誌
        self.assertEqual(self.mock_listing.WriteLine.call_count, 0)

    def test_batch_overlay_nx_ugraf_environment(self):
        """
        驗證在真實 NX 環境中（sys.executable 為 ugraf.exe，且當前進程無 PIL 庫）：
        1. is_valid_external_python 正確排除 ugraf.exe。
        2. resolve_python_runtime(require_pillow=True) 能探測到帶有 Pillow 的獨立 Python。
        3. batch_overlay_op_labels 能透過外部獨立 Python 順利完成標籤壓印。
        """
        from nx_cam_to_excel import is_valid_external_python, batch_overlay_op_labels
        from PIL import Image

        # 1. 驗證 ugraf.exe 絕不會被誤判為合法 Python
        self.assertFalse(is_valid_external_python(r"C:\Program Files\Siemens\NX2206\NXBIN\ugraf.exe"))
        self.assertFalse(is_valid_external_python(r"C:\NX\nxbin\python.exe"))

        # 2. 模擬 NX 宿主環境
        img = os.path.join(self.test_dir, "batch_ugraf_test.png")
        Image.new("RGB", (600, 400), (240, 240, 240)).save(img)

        with patch("sys.executable", r"C:\Program Files\Siemens\NX2206\NXBIN\ugraf.exe"):
            with patch.dict(sys.modules, {"PIL": None, "PIL.Image": None, "PIL.ImageDraw": None, "PIL.ImageFont": None}):
                ok = batch_overlay_op_labels({"OP_NX": img}, listing=self.mock_listing, force=True)
                self.assertTrue(ok)

        with Image.open(img) as im:
            self.assertEqual(im.getpixel((30, 30)), (0, 75, 135))


if __name__ == "__main__":
    unittest.main()


