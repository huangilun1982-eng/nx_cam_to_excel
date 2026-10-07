# -*- coding: utf-8 -*-
"""
組件單元測試：test_tool_conflict_detector.py
針對 tool_conflict_detector.py 進行覆蓋率驗證
"""
import unittest
from tool_conflict_detector import (
    parse_tool_numbers_from_str,
    normalize_tool_name_for_compare,
    safe_extract_float,
    detect_tool_conflicts_for_operations
)

class TestToolConflictDetector(unittest.TestCase):
    def test_parse_tool_numbers(self):
        self.assertEqual(parse_tool_numbers_from_str("T01"), [1])
        self.assertEqual(parse_tool_numbers_from_str("T12"), [12])
        self.assertEqual(parse_tool_numbers_from_str("5"), [5])
        self.assertEqual(parse_tool_numbers_from_str("T01~T03"), [1, 2, 3])
        self.assertEqual(parse_tool_numbers_from_str("T05~T05"), [5])
        self.assertEqual(parse_tool_numbers_from_str("-"), [])
        self.assertEqual(parse_tool_numbers_from_str(""), [])
        self.assertEqual(parse_tool_numbers_from_str(None), [])

    def test_normalize_tool_name(self):
        self.assertEqual(normalize_tool_name_for_compare("E4-T01"), "E4")
        self.assertEqual(normalize_tool_name_for_compare("E10_T02"), "E10")
        self.assertEqual(normalize_tool_name_for_compare("CC6"), "CC6")
        self.assertEqual(normalize_tool_name_for_compare("未指派"), "")
        self.assertEqual(normalize_tool_name_for_compare(None), "")

    def test_safe_extract_float(self):
        self.assertEqual(safe_extract_float("10.0"), 10.0)
        self.assertEqual(safe_extract_float("3.15"), 3.15)
        self.assertEqual(safe_extract_float("-"), 0.0)
        self.assertEqual(safe_extract_float("None"), 0.0)

    def test_normal_operations_no_conflict(self):
        ops = [
            {"tool_number": "T01", "raw_tool_name": "E10", "tool_diameter": "10.0", "note": "粗加工"},
            {"tool_number": "T02", "raw_tool_name": "E6", "tool_diameter": "6.0", "note": "半精加工"},
            {"tool_number": "T01", "raw_tool_name": "E10", "tool_diameter": "10.0", "note": "底部清角"} # 重複用同刀
        ]
        conflicts = detect_tool_conflicts_for_operations(ops)
        self.assertEqual(len(conflicts), 0)
        for op in ops:
            self.assertFalse(op.get("has_tool_conflict", False))
            self.assertNotIn("tool_conflict_msg", op)

    def test_conflict_detection(self):
        ops = [
            {"tool_number": "T01", "raw_tool_name": "E10", "tool_diameter": "10.0", "note": "粗加工"},
            {"tool_number": "T01", "raw_tool_name": "E6", "tool_diameter": "6.0", "note": "精加工"},
            {"tool_number": "T02", "raw_tool_name": "D4", "tool_diameter": "4.0", "note": "倒角"}
        ]
        conflicts = detect_tool_conflicts_for_operations(ops)
        self.assertEqual(len(conflicts), 1)
        self.assertTrue(ops[0].get("has_tool_conflict"))
        self.assertTrue(ops[1].get("has_tool_conflict"))
        self.assertFalse(ops[2].get("has_tool_conflict", False))

        self.assertIn("與 E6 同號", ops[0].get("tool_conflict_msg", ""))
        self.assertIn("與 E10 同號", ops[1].get("tool_conflict_msg", ""))

    def test_range_tool_conflict(self):
        ops = [
            {"tool_number": "T01~T02", "raw_tool_name": "E10", "tool_diameter": "10.0"},
            {"tool_number": "T02", "raw_tool_name": "E4", "tool_diameter": "4.0"}
        ]
        conflicts = detect_tool_conflicts_for_operations(ops)
        self.assertEqual(len(conflicts), 1)
        self.assertTrue(ops[0].get("has_tool_conflict"))
        self.assertTrue(ops[1].get("has_tool_conflict"))

if __name__ == "__main__":
    unittest.main()
