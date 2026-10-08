# -*- coding: utf-8 -*-
"""
單元測試：test_mcs_origin_detector.py
針對 mcs_origin_detector.py 的座標與方位解析、方案 2 階梯幾何提取進行單元測試
"""
import sys
import unittest
from unittest.mock import MagicMock

# 外部環境模擬 NXOpen 模組
for mod in ["NXOpen", "NXOpen.CAM", "NXOpen.UF"]:
    if mod not in sys.modules:
        sys.modules[mod] = MagicMock()

from mcs_origin_detector import (
    parse_fixture_offset_to_gcode,
    determine_axis_orientation,
    extract_target_bounding_box,
    resolve_stage_mcs_origin_string
)

class TestMCSOriginDetector(unittest.TestCase):
    def test_fixture_offset_conversion(self):
        self.assertEqual(parse_fixture_offset_to_gcode(0), "G54")
        self.assertEqual(parse_fixture_offset_to_gcode(1), "G54")
        self.assertEqual(parse_fixture_offset_to_gcode(2), "G55")
        self.assertEqual(parse_fixture_offset_to_gcode(3), "G56")
        self.assertEqual(parse_fixture_offset_to_gcode(4), "G57")
        self.assertEqual(parse_fixture_offset_to_gcode(5), "G58")
        self.assertEqual(parse_fixture_offset_to_gcode(6), "G59")
        self.assertEqual(parse_fixture_offset_to_gcode(54), "G54")
        self.assertEqual(parse_fixture_offset_to_gcode(55), "G55")
        self.assertEqual(parse_fixture_offset_to_gcode("invalid"), "G54")

    def test_xy_axis_orientations(self):
        # 零件 X: [-100, 100]
        # 中心點 0 -> MID
        self.assertEqual(determine_axis_orientation(0.0, -100.0, 100.0, is_z_axis=False), "MID")
        # 左側邊緣 -100 -> 0
        self.assertEqual(determine_axis_orientation(-100.0, -100.0, 100.0, is_z_axis=False), "0")
        # 右側邊緣 100 -> MAX
        self.assertEqual(determine_axis_orientation(100.0, -100.0, 100.0, is_z_axis=False), "MAX")

        # 零件以絕對 0 為基準邊: [0, 200]
        self.assertEqual(determine_axis_orientation(0.0, 0.0, 200.0, is_z_axis=False), "0")
        self.assertEqual(determine_axis_orientation(100.0, 0.0, 200.0, is_z_axis=False), "MID")
        self.assertEqual(determine_axis_orientation(200.0, 0.0, 200.0, is_z_axis=False), "MAX")

    def test_z_axis_orientations(self):
        # 1. 零件頂面為絕對 0: [-50, 0]，原點在頂面 0 -> TOP
        self.assertEqual(determine_axis_orientation(0.0, -50.0, 0.0, is_z_axis=True), "TOP")
        # 2. 零件頂面為 0，原點高於頂面 (預留切削量 0.5mm) -> TOP
        self.assertEqual(determine_axis_orientation(0.5, -50.0, 0.0, is_z_axis=True), "TOP")
        # 3. 素材頂面為 0.5，原點在工件頂面 0 (靠近頂面 0.5mm，遠離中心 24.75mm) -> TOP
        self.assertEqual(determine_axis_orientation(0.0, -50.0, 0.5, is_z_axis=True), "TOP")
        # 4. 原點在頂面偏下一點點台階面 -0.2mm (靠近頂面) -> TOP
        self.assertEqual(determine_axis_orientation(-0.2, -50.0, 0.0, is_z_axis=True), "TOP")

        # 5. 零件底面為絕對 0: [0, 50]，原點在底面 0 -> 0
        self.assertEqual(determine_axis_orientation(0.0, 0.0, 50.0, is_z_axis=True), "0")
        # 6. 原點低於底面 (台面碰刀 -5mm) -> 0
        self.assertEqual(determine_axis_orientation(-5.0, 0.0, 50.0, is_z_axis=True), "0")

        # 7. 中心厚度分中: [-25, 25]，原點在 0 -> MID
        self.assertEqual(determine_axis_orientation(0.0, -25.0, 25.0, is_z_axis=True), "MID")
        # 8. 中心偏厚度: [-50, 0]，中心 -25 -> MID
        self.assertEqual(determine_axis_orientation(-25.0, -50.0, 0.0, is_z_axis=True), "MID")

        # 9. 零件厚度很薄 (如 1mm 薄板): [-1.0, 0.0]，原點在 0.0 -> TOP
        self.assertEqual(determine_axis_orientation(0.0, -1.0, 0.0, is_z_axis=True), "TOP")
        # 10. 薄板底面: [-1.0, 0.0]，原點在 -1.0 -> 0
        self.assertEqual(determine_axis_orientation(-1.0, -1.0, 0.0, is_z_axis=True), "0")

    def test_extract_target_bounding_box_scheme2_cam_priority(self):
        """
        方案 2 階梯 1 測試：CAM 工件/素材指定幾何優先
        """
        mock_wp = MagicMock()
        mock_setup = MagicMock()
        mock_uf = MagicMock()

        # 模擬 WORKPIECE 節點
        mock_node = MagicMock()
        mock_node.Name = "WORKPIECE"
        mock_node.GetMembers.return_value = []
        mock_root = MagicMock()
        mock_root.Name = "GEOMETRY_ROOT"
        mock_root.GetMembers.return_value = [mock_node]
        mock_setup.GetRoot.return_value = mock_root

        mock_builder = MagicMock()
        mock_setup.CAMGroupCollection.CreateMillGeomBuilder.return_value = mock_builder

        mock_cam_body = MagicMock()
        mock_cam_body.Tag = 1001

        mock_gset = MagicMock()
        mock_gset.GetItems.return_value = [mock_cam_body]
        mock_builder.PartGeometry.GeometryList.Length = 1
        mock_builder.PartGeometry.GeometryList.FindItem.return_value = mock_gset
        mock_builder.BlankGeometry.GeometryList.Length = 0

        # AskBoundingBox 回傳 [-50, -50, -20, 50, 50, 0]
        mock_uf.ModlGeneral.AskBoundingBox.return_value = [-50.0, -50.0, -20.0, 50.0, 50.0, 0.0]

        found, b_min, b_max, source = extract_target_bounding_box(
            work_part=mock_wp, cam_setup=mock_setup, uf_session=mock_uf
        )

        self.assertTrue(found)
        self.assertEqual(source, "CAM_GEOMETRY")
        self.assertEqual(b_min, [-50.0, -50.0, -20.0])
        self.assertEqual(b_max, [50.0, 50.0, 0.0])

    def test_extract_target_bounding_box_scheme2_display_objects(self):
        """
        方案 2 階梯 2 測試：CAM 未指定時，只抓圖面顯示物件 (過濾圖層 4 與 IsBlanked)
        """
        mock_wp = MagicMock()
        mock_setup = MagicMock()
        mock_uf = MagicMock()

        # 讓 CAM 搜尋不到任何實體
        mock_root = MagicMock()
        mock_root.Name = "GEOMETRY_ROOT"
        mock_root.GetMembers.return_value = []
        mock_setup.GetRoot.return_value = mock_root

        # 模擬 3 個實體：
        # b1: 圖層 1 (正常顯示)
        # b2: 圖層 4 (隱藏圖層)
        # b3: 圖層 1 但 IsBlanked = True (隱藏)
        b1 = MagicMock()
        b1.Layer = 1
        b1.IsBlanked = False
        b1.Tag = 2001

        b2 = MagicMock()
        b2.Layer = 4
        b2.IsBlanked = False
        b2.Tag = 2002

        b3 = MagicMock()
        b3.Layer = 1
        b3.IsBlanked = True
        b3.Tag = 2003

        mock_wp.Bodies = [b1, b2, b3]

        def ask_status_side_effect(layer):
            if layer == 4:
                return 4  # 隱藏圖層
            return 1      # 正常工作圖層

        mock_uf.Layer.AskStatus.side_effect = ask_status_side_effect

        def ask_bbox_side_effect(tag):
            if tag == 2001:
                return [-10.0, -10.0, -5.0, 10.0, 10.0, 0.0]
            elif tag == 2002:
                # 若抓到 b2 則範圍會被擴大到 999
                return [-999.0, -999.0, -999.0, 999.0, 999.0, 999.0]
            elif tag == 2003:
                return [-888.0, -888.0, -888.0, 888.0, 888.0, 888.0]
            return [0, 0, 0, 0, 0, 0]

        mock_uf.ModlGeneral.AskBoundingBox.side_effect = ask_bbox_side_effect

        found, b_min, b_max, source = extract_target_bounding_box(
            work_part=mock_wp, cam_setup=mock_setup, uf_session=mock_uf
        )

        self.assertTrue(found)
        self.assertEqual(source, "DISPLAY_BODIES")
        # 驗證只有 b1 被納入，b2 與 b3 被成功過濾排除！
        self.assertEqual(b_min, [-10.0, -10.0, -5.0])
        self.assertEqual(b_max, [10.0, 10.0, 0.0])

    def test_extract_target_bounding_box_scheme2_all_bodies_fallback(self):
        """
        方案 2 階梯 3 測試：當顯示物件為空時，回退全零件實體保底
        """
        mock_wp = MagicMock()
        mock_setup = MagicMock()
        mock_uf = MagicMock()

        mock_root = MagicMock()
        mock_root.Name = "GEOMETRY_ROOT"
        mock_root.GetMembers.return_value = []
        mock_setup.GetRoot.return_value = mock_root

        # 唯一實體但處於隱藏圖層 4
        b1 = MagicMock()
        b1.Layer = 4
        b1.IsBlanked = False
        b1.Tag = 3001
        mock_wp.Bodies = [b1]

        mock_uf.Layer.AskStatus.return_value = 4
        mock_uf.ModlGeneral.AskBoundingBox.return_value = [-30.0, -30.0, -10.0, 30.0, 30.0, 0.0]

        found, b_min, b_max, source = extract_target_bounding_box(
            work_part=mock_wp, cam_setup=mock_setup, uf_session=mock_uf
        )

        self.assertTrue(found)
        self.assertEqual(source, "ALL_BODIES")
        self.assertEqual(b_min, [-30.0, -30.0, -10.0])
        self.assertEqual(b_max, [30.0, 30.0, 0.0])

if __name__ == "__main__":
    unittest.main()
