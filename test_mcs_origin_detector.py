# -*- coding: utf-8 -*-
"""
單元測試：test_mcs_origin_detector.py
針對 mcs_origin_detector.py 的座標與方位解析進行單元測試
"""
import unittest
from mcs_origin_detector import (
    parse_fixture_offset_to_gcode,
    determine_axis_orientation
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
        # 零件 Z: [-50, 0]
        # 頂面 0 -> TOP
        self.assertEqual(determine_axis_orientation(0.0, -50.0, 0.0, is_z_axis=True), "TOP")
        # 底面 -50 -> 0
        self.assertEqual(determine_axis_orientation(-50.0, -50.0, 0.0, is_z_axis=True), "0")
        # 中間 -25 -> MID
        self.assertEqual(determine_axis_orientation(-25.0, -50.0, 0.0, is_z_axis=True), "MID")

        # 零件 Z: [-20, 20]
        self.assertEqual(determine_axis_orientation(20.0, -20.0, 20.0, is_z_axis=True), "TOP")
        self.assertEqual(determine_axis_orientation(0.0, -20.0, 20.0, is_z_axis=True), "MID")
        self.assertEqual(determine_axis_orientation(-20.0, -20.0, 20.0, is_z_axis=True), "0")

if __name__ == "__main__":
    unittest.main()
