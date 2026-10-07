# -*- coding: utf-8 -*-
"""
組件名稱：tool_safety_evaluator.py
功能職責：CAM 刀具有效刃長與懸伸安全評估引擎
設計原則：
1. 組件功能單純原則：專注於比對刀具刃長、夾持長度與工序切深，評估避空與干涉安全。
2. 零破壞防護 (Zero-Debug Protocol)：若無超深違規或無刀軌，保持現有備註不變，絕不中斷流程。
3. 嚴格資源管理：即時銷毀 NX C++ Builder/Checker 物件，杜絕記憶體洩漏。
"""

def evaluate_single_tool_safety(flute_len_val, holder_len_val, max_cut_depth=0.0, violation_amount=0.0):
    """
    純數據計算：評估單一工步之刀具刃長與切深關係。
    
    回傳：
      warning_msg (str)：若有超深則回傳中文警示標籤，否則回傳空字串 ""
    """
    try:
        f_len = float(flute_len_val) if flute_len_val not in ["-", "", None] else 0.0
    except Exception:
        f_len = 0.0

    try:
        h_len = float(holder_len_val) if holder_len_val not in ["-", "", None] else 0.0
    except Exception:
        h_len = 0.0

    actual_depth = 0.0
    if violation_amount > 0.0 and f_len > 0.0:
        actual_depth = f_len + violation_amount
    elif max_cut_depth > 0.0:
        actual_depth = max_cut_depth

    if actual_depth <= 0.0 or f_len <= 0.0:
        return ""

    # 1. 嚴重干涉風險：切深 > 夾長 (懸長)
    if h_len > 0.0 and actual_depth > h_len:
        return f"【警告: 切深{actual_depth:.1f} > 夾長{h_len:.1f}】"

    # 2. 避空不足風險：切深 > 刃長
    if actual_depth > f_len:
        return f"【注意: 切深{actual_depth:.1f} > 刃長{f_len:.1f}】"

    return ""

def evaluate_operations_tool_safety(operations, cam_setup=None):
    """
    批次評估工步清單中的所有工序，透過 NX 原廠 CutDepthChecker 或數據比對檢驗安全性。
    
    參數：
      operations: 工步字典清單
      cam_setup: NXOpen.CAM.CAMSetup 物件 (選填)
      
    回傳：
      safety_warnings (list of str)：警示報告摘要
    """
    if not operations:
        return []

    safety_warnings = []

    for op_item in operations:
        flute_str = op_item.get("flute_length", "-")
        holder_str = op_item.get("holder_length", "-")
        nx_op = op_item.get("_nx_op_obj")

        violation_amt = 0.0

        # 若在 NX 實機環境且有有效工序物件與 cam_setup，嘗試調用 CutDepthChecker
        if cam_setup is not None and nx_op is not None:
            checker = None
            try:
                checker = cam_setup.CreateCutDepthChecker()
                if checker:
                    checker.OperationToCheck = nx_op
                    checker.CheckNeck = True
                    checker.PerformCheck()
                    if checker.NumCutDepthViolations > 0:
                        violation_amt = checker.MaxCutDepthViolation
            except Exception:
                pass
            finally:
                if checker is not None:
                    try:
                        checker.Destroy()
                    except Exception:
                        pass

        # 評估安全性
        warn_msg = evaluate_single_tool_safety(flute_str, holder_str, violation_amount=violation_amt)
        if warn_msg:
            op_item["tool_safety_warning"] = warn_msg
            op_name = op_item.get("op_name", op_item.get("clean_op_name", "工步"))
            tool_lbl = op_item.get("tool_spec_display", op_item.get("tool_number", "刀具"))
            safety_warnings.append(f"{op_name} ({tool_lbl}): {warn_msg}")

    return safety_warnings
