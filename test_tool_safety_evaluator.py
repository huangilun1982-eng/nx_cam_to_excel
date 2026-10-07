# -*- coding: utf-8 -*-
"""
單元測試：test_tool_safety_evaluator.py
驗證刀具有效刃長與懸伸安全評估組件
"""
import unittest
from tool_safety_evaluator import (
    evaluate_single_tool_safety,
    evaluate_operations_tool_safety
)

class TestToolSafetyEvaluator(unittest.TestCase):
    def test_evaluate_single_tool_safe(self):
        # 正常切削 (切深 15 < 刃長 20)
        warn = evaluate_single_tool_safety("20.0", "30.0", max_cut_depth=15.0)
        self.assertEqual(warn, "")

    def test_evaluate_single_tool_flute_exceeded(self):
        # 避空不足 (刃長 20，違規超深 8mm -> 實際切深 28mm)
        warn = evaluate_single_tool_safety("20.0", "35.0", violation_amount=8.0)
        self.assertEqual(warn, "【注意: 切深28.0 > 刃長20.0】")

    def test_evaluate_single_tool_holder_exceeded(self):
        # 嚴重干涉 (刃長 20，夾長 25，違規超深 10mm -> 實際切深 30mm > 夾長 25mm)
        warn = evaluate_single_tool_safety("20.0", "25.0", violation_amount=10.0)
        self.assertEqual(warn, "【警告: 切深30.0 > 夾長25.0】")

    def test_evaluate_operations_batch(self):
        ops = [
            {"op_name": "ROUGH", "flute_length": "20.0", "holder_length": "35.0"},
            {"op_name": "FINISH", "flute_length": "15.0", "holder_length": "25.0"}
        ]
        # 無 cam_setup 時安全靜態評估
        res = evaluate_operations_tool_safety(ops, cam_setup=None)
        self.assertEqual(len(res), 0)
        self.assertNotIn("tool_safety_warning", ops[0])

if __name__ == "__main__":
    unittest.main()
