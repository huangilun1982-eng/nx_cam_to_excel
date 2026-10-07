# -*- coding: utf-8 -*-
"""
組件名稱：tool_conflict_detector.py
功能職責：CAM 刀具 T 號重複性衝突檢驗引擎
設計原則：
1. 組件功能單純原則：專注於刀號解析與工步刀具規格一致性比對，不涉及 UI 或檔案寫入。
2. 零破壞防護：純資料運算，若無衝突則 100% 保持原始資料完整性。
3. 高可靠度防呆：支援單一刀號 (T01)、範圍刀號 (T01~T03)、無效刀號防禦及同刀多次使用白名單。
"""

import re

def parse_tool_numbers_from_str(val):
    """
    從刀號字串或數值中安全解析出所有涵蓋的整數刀號清單。
    
    支援格式：
      - "T01" / "1" / 1 -> [1]
      - "T01~T03" / "1~3" -> [1, 2, 3]
      - "-" / "" / None -> []
    """
    if val is None:
        return []
    s = str(val).strip().upper()
    if not s or s in ["-", "NONE"]:
        return []

    # 1. 支援範圍格式 (例如連續刀號合併後的 T01~T03 或 1~3)
    range_match = re.match(r'^[T]?(\d+)\s*~\s*[T]?(\d+)$', s)
    if range_match:
        try:
            start_n = int(range_match.group(1))
            end_n = int(range_match.group(2))
            if start_n <= end_n:
                return list(range(start_n, end_n + 1))
            else:
                return [start_n, end_n]
        except Exception:
            return []

    # 2. 單一刀號格式 (例如 T01, T2, 5)
    m = re.search(r'\d+', s)
    if m:
        try:
            return [int(m.group())]
        except Exception:
            return []

    return []

def normalize_tool_name_for_compare(name):
    """
    標準化刀具名稱以供比對：
    移除尾部常見的工步/刀號後綴 (例如 E4-T01, E4_T01 -> E4)
    """
    if not name or str(name).strip() in ["未指派", "None", "-", ""]:
        return ""
    clean = str(name).strip()
    clean = re.sub(r'[-_]T\d+$', '', clean, flags=re.IGNORECASE)
    return clean.strip()

def safe_extract_float(val):
    """
    安全提取浮點數直徑或尺寸數值，解析失敗回傳 0.0
    """
    if not val or str(val).strip() in ["-", "", "None"]:
        return 0.0
    try:
        m = re.search(r'[-+]?\d*\.?\d+', str(val))
        if m:
            return float(m.group())
    except Exception:
        pass
    return 0.0

def detect_tool_conflicts_for_operations(operations):
    """
    針對一組工步（通常為同一個工段 Stage 下的所有工序/工步）執行 T 號重複性衝突檢驗。

    比對準則：
      - 機台同工段內，每個 T 號應對應唯一規格之刀具。
      - 若相同 T 號指派給不同刀名或不同直徑，視為【T 號衝突】。
      - 若相同 T 號多次出現但幾何與刀名相同（同把刀重複切削不同特徵），屬正常製程，不報衝突。

    回傳值：
      conflicts_summary (list of str)：衝突摘要文字清單，供記錄與視窗警示提示。
      同時在衝突的 operation 字典中直接更新：
        op["has_tool_conflict"] = True
        op["tool_conflict_msg"] = "【衝突: 與 [xxx] 同號】"
    """
    if not operations:
        return []

    # 1. 建立 T 號 -> 所有使用此 T 號之工步項目的關聯表
    tool_map = {}

    for idx, op in enumerate(operations):
        t_str = op.get("tool_number", "-")
        t_ints = parse_tool_numbers_from_str(t_str)
        if not t_ints:
            continue

        raw_name = op.get("raw_tool_name", op.get("tool_name", ""))
        clean_name = normalize_tool_name_for_compare(raw_name)
        dia = safe_extract_float(op.get("tool_diameter", 0))

        # 顯示標籤 (例如 E10 或 D10.0)
        spec_display = op.get("tool_spec_display")
        if not spec_display or spec_display in ["-", "None", ""]:
            spec_display = clean_name if clean_name else (f"D{dia:.2f}".rstrip('0').rstrip('.') if dia > 0 else "未知刀具")

        # 刀具唯一識別碼：(乾淨名稱大寫, 四捨五入直徑)
        identity_key = (clean_name.upper(), round(dia, 2))

        for t_num in t_ints:
            if t_num not in tool_map:
                tool_map[t_num] = []
            tool_map[t_num].append({
                "op_idx": idx,
                "op": op,
                "t_num": t_num,
                "clean_name": clean_name,
                "dia": round(dia, 2),
                "identity_key": identity_key,
                "label": spec_display
            })

    conflicts_summary = []

    # 2. 檢驗每個 T 號是否存在 2 種以上不同規格
    for t_num, entries in tool_map.items():
        distinct_identities = {}
        for entry in entries:
            ikey = entry["identity_key"]
            if ikey not in distinct_identities:
                distinct_identities[ikey] = entry["label"]

        if len(distinct_identities) > 1:
            # 發生衝突！
            all_labels = list(distinct_identities.values())
            summary_msg = f"刀號 T{t_num:02d} 重複指派給不同刀具規格: {' vs '.join(all_labels)}"
            conflicts_summary.append(summary_msg)

            # 為受影響的工步註記衝突狀態與說明
            for entry in entries:
                target_op = entry["op"]
                target_op["has_tool_conflict"] = True

                # 找出與當前工步不同的其他衝突刀具名稱
                other_labels = [lbl for ikey, lbl in distinct_identities.items() if ikey != entry["identity_key"]]
                conflict_desc = f"【衝突: 與 {'/'.join(other_labels)} 同號】"

                existing_msg = target_op.get("tool_conflict_msg", "")
                if existing_msg:
                    if conflict_desc not in existing_msg:
                        target_op["tool_conflict_msg"] = f"{existing_msg} {conflict_desc}"
                else:
                    target_op["tool_conflict_msg"] = conflict_desc

    return conflicts_summary
