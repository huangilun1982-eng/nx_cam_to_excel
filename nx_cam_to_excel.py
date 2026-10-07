# -*- coding: utf-8 -*-
"""
Siemens NX CAM 加工工序單自動化匯出工具 (工業級分頁與智慧辨識版)
功能：
  1. 支援 CAM 導覽器群組與工序選取，智慧識別資訊資料夾 (材料/料號/孔數/厚度)
  2. 自動排除 PROGRAM、未用項 等容器名稱
  3. 支援 -P / -p 裝夾翻面強制分頁邏輯
  4. 支援每頁 10 列自動換頁與群組完整性保證 (Keep Group Together)
  5. 自動複製 Excel 母版生成「第1頁」、「第2頁」...，動態填寫 1-N 頁碼與表頭
  6. 刀徑整合刀具名稱型號、刀號標準化 T01、夾長防撞安全護欄運算
  7. 序號與程式檔名規範：同群組首道工序填寫群組檔名，後續工步留空；序號為每頁有效工步連續流水號 (1, 2, 3...)，空行不計序號
  8. 刀號智慧合併機制：支援同刀號切削時間自動累計、連續刀號同規格合併 (如 T01~T02) 與防撞長度最大值保護
  9. 表頭欄位精確填入：圖名為檔案名稱；圖號若包含 14 碼編號則填入、否則留空；尺寸由素材大小決定、無設定則留空
"""

import os
import sys
import re
import json
import datetime
import subprocess
import tempfile
import glob
import shutil
try:
    import winreg
except ImportError:
    winreg = None
import NXOpen
import NXOpen.CAM
import NXOpen.UF
try:
    import NXOpen.Gateway
except Exception:
    pass

# ==================== 設定區 (參考 ExcelTool 規則) ====================
ROWS_PER_PAGE = 10                  # ShopDoc_Template.xlsx 每頁工步上限 (Row 7 ~ Row 16)
EXCLUDE_KEYWORDS = ["NC_PROGRAM", "未用項"]  # 排除群組關鍵字
FORCE_PAGE_SUFFIXES = ["-M", "-m", "_M", "_m", "-P", "-p", "_P", "_p"] # 強制分頁後綴 (支援 M/P 裝夾翻面)

# 資訊資料夾前綴與表頭欄位對照
INFO_PREFIX_RULES = {
    "尺寸": "size",
    "素材尺寸": "size",
    "素材": "size",
    "材料": "material",
    "材質": "material",
    "料號": "part_number",
    "工單": "part_number",
    "厚度": "thickness",
    "孔數": "holes",
    "數量": "holes",
    "顏色": "color"
}

# ==================== CAM 物件收集與解析模組 ====================

def clean_program_name(name):
    """
    程式檔名清洗規則：移除結尾的 -Txx 或 _Txx (例如 D6.0-T01 -> D6.0)
    """
    if not name:
        return ""
    clean = str(name).strip()
    clean = re.sub(r'[-_]T\d+$', '', clean, flags=re.IGNORECASE)
    return clean

def is_excluded_group_name(name):
    """
    檢查群組名稱是否在排除清單中 (如未用項)
    """
    if not name:
        return False
    clean = str(name).strip()
    for kw in EXCLUDE_KEYWORDS:
        if kw in clean:
            return True
    return False

def check_and_extract_info(name, extracted_info):
    """
    檢查是否為資訊資料夾 (如 材料AL6061、料號75612960)
    若符合則提取後方字串並返回 True
    """
    if not name:
        return False
    clean = str(name).strip()
    for prefix, key in INFO_PREFIX_RULES.items():
        if clean.startswith(prefix):
            val = clean[len(prefix):].strip(" :_-=")
            if val:
                extracted_info[key] = val
            return True
    return False

def extract_stage_key(name):
    """
    從群組或工序名稱中精確解析工段標籤 (以 M1, M2, M3... 為核心，亦支援 OP10 與舊版 P1, P2)
    若無明確工段標籤則回傳 None。
    安全機制：限制序號為 1~49，排除刀柄懸伸長度 (如 P150, P200) 與一般規格字串。
    """
    if not name or str(name).strip().upper() in ["DEFAULT", "PROGRAM", "NC_PROGRAM", "NONE", "-"]:
        return None

    clean = str(name).strip()

    # 1. 匹配中綴/後綴 [-_](M數字, OP數字, P數字, M, P)
    # 限制 M/P 數字最多 2 位數且 < 50，排除大於等於 50 的刀柄懸伸規格 (如 P150, P200)
    m_m = re.search(r'[-_](M\d{1,2}|OP\d{1,2}|P\d{1,2}|M|m|P|p)(?:[-_\s]|$)', clean, re.IGNORECASE)
    if m_m:
        val = m_m.group(1).upper()
        if val in ["M", "P"]:
            return "M1"
        if val.startswith("P") and val[1:].isdigit():
            p_num = int(val[1:])
            if p_num < 50:
                return f"M{p_num}"
            return None
        if val.startswith("M") and val[1:].isdigit():
            m_num = int(val[1:])
            if m_num < 50:
                return f"M{m_num}"
            return None
        return val

    # 2. 匹配開頭的 M1, M2, OP10, OP20, P1, P2
    m_op = re.match(r'^(M\d{1,2}|OP\d{1,2}|P\d{1,2})(?:[-_\s]|$)', clean, re.IGNORECASE)
    if m_op:
        val = m_op.group(1).upper()
        if val.startswith("P") and val[1:].isdigit():
            p_num = int(val[1:])
            if p_num < 50:
                return f"M{p_num}"
            return None
        if val.startswith("M") and val[1:].isdigit():
            m_num = int(val[1:])
            if m_num < 50:
                return f"M{m_num}"
            return None
        return val

    return None

def build_cam_tree_map(root_group, extracted_info=None):
    """
    從 ProgramOrder 樹根節點建立全域父子與頂層母資料夾映射表 (以物件 Tag 為 Key)
    並自動全域掃描提取資訊資料夾 (ExcelTool 模式：材料/尺寸/厚度/料號等)
    Tag -> {"top_parent": str, "stage": str, "direct_group": str}
    """
    node_info = {}

    def traverse(node, current_top=None, current_stage=None, direct_parent_group=None):
        if not node:
            return

        node_name = getattr(node, "Name", "").strip()
        is_grp = isinstance(node, NXOpen.CAM.NCGroup)
        stg = extract_stage_key(node_name)

        next_top = current_top
        next_stage = current_stage
        next_direct = direct_parent_group

        if is_grp:
            # 1. 檢查並自動提取資訊資料夾 (ExcelTool 模式：材料、尺寸、料號、厚度等)
            if extracted_info is not None and check_and_extract_info(node_name, extracted_info):
                # 這是資訊資料夾，提取資訊後不當作加工程式群組，直接返回
                return

            # 2. 排除群組判定
            if is_excluded_group_name(node_name):
                return

            if node_name.upper() not in ["PROGRAM", "NC_PROGRAM", "未用項"]:
                if current_top is None or stg:
                    # 這是頂層父資料夾！例如 MK-1-M1, MK-2-M2
                    next_top = node_name
                    next_stage = stg or current_stage or "通用工段"
                next_direct = node_name

        tag = getattr(node, "Tag", None)
        if tag:
            node_info[tag] = {
                "top_parent": next_top or "DEFAULT",
                "stage": next_stage or "通用工段",
                "direct_group": next_direct or "DEFAULT"
            }

        if is_grp:
            try:
                for child in node.GetMembers():
                    traverse(child, next_top, next_stage, next_direct)
            except Exception:
                pass

    try:
        traverse(root_group)
    except Exception:
        pass
    return node_info

def resolve_stage_and_parent_program(node, current_stage=None, current_parent=None, tree_map=None):
    """
    沿著 CAM 導覽器樹狀結構向上回溯，精準鎖定所屬 (工段標籤, 母程式父資料夾名稱, 直接群組名稱)
    核心原則：
    1. 優先查全域樹地圖 (保證 100% 精確)
    2. 工序 (Operation) 絕非資料夾，溯源時直接跳過，只收集 NCGroup
    3. 父資料夾為頂層 NCGroup (緊鄰 PROGRAM 下方) 或具備工段標籤 (如 MK-1-M1) 者
    """
    tag = getattr(node, "Tag", None)
    if tree_map and tag and tag in tree_map:
        info = tree_map[tag]
        return info["stage"], info["top_parent"], info["direct_group"]

    curr = node
    # 若當前節點是工序，必須先跳到其所屬群組，絕不能將工序名當成程式群組
    if isinstance(curr, NXOpen.CAM.Operation):
        try:
            curr = curr.GetParent(NXOpen.CAM.CAMSetup.View.ProgramOrder)
        except Exception:
            curr = None

    ancestor_groups = []
    while curr is not None:
        name = getattr(curr, "Name", "").strip()
        is_grp = isinstance(curr, NXOpen.CAM.NCGroup)
        if is_grp and name.upper() not in ["PROGRAM", "NC_PROGRAM", "DEFAULT", "未用項", ""]:
            ancestor_groups.append((curr, name))
        try:
            curr = curr.GetParent(NXOpen.CAM.CAMSetup.View.ProgramOrder)
        except Exception:
            break

    if not ancestor_groups:
        fallback_stage = current_stage or "通用工段"
        fallback_parent = current_parent or "DEFAULT"
        return fallback_stage, fallback_parent, "DEFAULT"

    # ancestor_groups 由底至頂：[子資料夾, 父資料夾]
    direct_group = ancestor_groups[0][1]

    # 尋找帶有工段標籤之群組 (例如 MK-1-M1)
    stage_found = None
    stage_prog_name = None
    for _, grp_name in reversed(ancestor_groups):
        stg = extract_stage_key(grp_name)
        if stg:
            stage_found = stg
            stage_prog_name = grp_name
            break

    # 最頂層的有效群組 (緊鄰 PROGRAM 之下)
    top_group_name = ancestor_groups[-1][1]

    final_stage = stage_found or current_stage or "通用工段"
    final_prog = stage_prog_name or top_group_name or current_parent or "DEFAULT"
    return final_stage, final_prog, direct_group

def collect_cam_hierarchy(obj, target_chunks, extracted_info, current_group_name="", current_stage=None, parent_program=None, tree_map=None, visited_tags=None, directly_selected_tags=None):
    """
    遞迴走訪 CAM 物件，依據群組歸納為 Chunk，並過濾排除群組與提取資訊群組。
    各子資料夾維持獨立 Chunk (不跨子資料夾合併)，並自動溯源鎖定所屬「母程式父資料夾」與工段標籤。
    具備工序自動去重防護，並標記直接選取的工序。
    """
    if isinstance(obj, NXOpen.CAM.Operation):
        tag = getattr(obj, "Tag", None)
        # 工序自動去重防護 (避免重複/重疊選取導致工步重複輸出)
        if visited_tags is not None and tag:
            if tag in visited_tags:
                return
            visited_tags.add(tag)

        # 標記是否為使用者直接點選之工序
        is_direct = False
        if directly_selected_tags is not None and tag:
            is_direct = (tag in directly_selected_tags)
        try:
            setattr(obj, "_is_direct_op", is_direct)
        except Exception:
            pass

        # 決定所屬之直接子群組、工段與母程式父資料夾
        if tree_map and tag and tag in tree_map:
            t_stage = tree_map[tag]["stage"]
            t_top = tree_map[tag]["top_parent"]
            t_direct = tree_map[tag]["direct_group"]
        else:
            t_stage, t_top, t_direct = resolve_stage_and_parent_program(obj, current_stage, parent_program, tree_map)

        op_group = current_group_name or t_direct or "DEFAULT"
        stage = current_stage or t_stage or "通用工段"
        top_prog = parent_program or t_top or "DEFAULT"

        if not target_chunks or target_chunks[-1]["group_name"] != op_group or target_chunks[-1]["stage"] != stage:
            target_chunks.append({
                "group_name": op_group,
                "parent_program": top_prog,
                "stage": stage,
                "operations": [obj]
            })
        else:
            target_chunks[-1]["operations"].append(obj)

    elif isinstance(obj, NXOpen.CAM.NCGroup):
        grp_name = obj.Name.strip()

        # 1. 檢查是否為資訊資料夾 (如: 材料AL6061)
        if check_and_extract_info(grp_name, extracted_info):
            return

        # 2. 檢查是否在排除清單中 (如: 未用項)
        if is_excluded_group_name(grp_name):
            return

        # 3. 溯源/繼承工段與母程式群組
        t_stage, t_top, _ = resolve_stage_and_parent_program(obj, current_stage, parent_program, tree_map)
        detected_stage = extract_stage_key(grp_name)

        if grp_name.upper() in ["PROGRAM", "NC_PROGRAM"]:
            next_group_name = current_group_name
            next_stage = current_stage
            next_parent = parent_program
        elif detected_stage or not parent_program:
            # 程式群組起點 (如選取之頂層母資料夾 MK-1-M1 或獨立選取之子資料夾 E20-125L，程式檔名 100% 與選取名稱一致)
            next_group_name = grp_name
            next_stage = detected_stage if detected_stage else t_stage
            next_parent = grp_name
        else:
            # 已經在某個母群組內部之子資料夾 (如 E20-125L)：
            # current_group_name 設為該子資料夾以維持各子資料夾獨立 Chunk (不跨組合併)，母程式父資料夾名稱向下維持！
            next_group_name = grp_name
            next_stage = current_stage
            next_parent = parent_program

        members = obj.GetMembers()
        for member in members:
            collect_cam_hierarchy(member, target_chunks, extracted_info, next_group_name, next_stage, next_parent, tree_map, visited_tags, directly_selected_tags)

# ==================== 表頭資訊萃取模組 (圖號/圖名/素材尺寸) ====================

def extract_14_digit_drawing_number(name):
    """
    從檔案名稱中提取 14 碼編號 (若當中有 14 碼編號則為圖號，若無則留空)
    """
    if not name:
        return ""
    m = re.search(r'(?<![A-Za-z0-9])([A-Za-z0-9]{14})(?![A-Za-z0-9])', str(name))
    return m.group(1) if m else ""

def extract_blank_size_from_group_name(group_name):
    """
    從程式群組名稱中探測素材規格 (例如 160x160x10-P1 -> 160x160x10)
    """
    if not group_name:
        return ""
    m = re.search(r'(\d+(?:\.\d+)?)\s*[xX*]\s*(\d+(?:\.\d+)?)\s*[xX*]\s*(\d+(?:\.\d+)?)', str(group_name))
    if m:
        def fmt_s(s):
            try:
                f = float(s)
                return str(int(f)) if f.is_integer() else s
            except Exception:
                return s
        return f"{fmt_s(m.group(1))}x{fmt_s(m.group(2))}x{fmt_s(m.group(3))}"
    return ""

def extract_blank_size_from_cam(work_part, uf_session):
    """
    嘗試從 NX CAM WORKPIECE 設定的素材幾何 (Blank) 讀取其邊界尺寸 (長x寬x高)
    """
    try:
        cam_setup = work_part.CAMSetup
        if not cam_setup:
            return ""

        geom_root = cam_setup.GetRoot(NXOpen.CAM.CAMSetup.View.Geometry)
        if not geom_root:
            return ""

        nodes_to_check = []
        def traverse_geom(node):
            if not node:
                return
            nodes_to_check.append(node)
            try:
                for member in node.GetMembers():
                    if isinstance(member, NXOpen.CAM.NCGroup):
                        traverse_geom(member)
            except Exception:
                pass

        traverse_geom(geom_root)

        for g_node in nodes_to_check:
            try:
                name_upper = g_node.Name.upper()
                if "WORKPIECE" in name_upper or "BLANK" in name_upper:
                    box = uf_session.Bnd.AskBox(g_node.Tag)
                    if box and len(box) >= 6:
                        lx, ly, lz = box[3], box[4], box[5]
                        if lx > 0.5 and ly > 0.5 and lz > 0.5:
                            def fmt_val(v):
                                return str(int(round(v))) if abs(v - round(v)) < 0.05 else f"{v:.1f}"
                            return f"{fmt_val(lx)}x{fmt_val(ly)}x{fmt_val(lz)}"
            except Exception:
                pass
    except Exception:
        pass
    return ""

def determine_blank_size(work_part, uf_session, raw_chunks, extracted_info):
    """
    尺寸應該由取得設定的素材大小決定，若無資訊則留空 ("")
    優先順序 (完全整合 ExcelTool 模式與 NX 原生幾何讀取)：
    1. 資訊資料夾明確定義 (ExcelTool 經典模式)：
       - 優先讀取「尺寸」/「素材」前綴資料夾 (如 尺寸160x160x10、素材160*160*10)
       - 相容讀取 ExcelTool「材料」/「材質」前綴資料夾 (如 材料160x160x10、材料AL6061 160x160x10)
    2. 選取的程式群組名稱中內含的素材規格 (如 160x160x10-P1)
    3. NX CAM WORKPIECE 設定的素材幾何邊界 (原生 3D AskBox 量測)
    4. 若無資訊則留空 ("")
    """
    # 1. 資訊資料夾 (ExcelTool 模式)
    if "size" in extracted_info and extracted_info["size"]:
        raw_sz = extracted_info["size"]
        clean_sz = extract_blank_size_from_group_name(raw_sz)
        return clean_sz if clean_sz else raw_sz

    if "material" in extracted_info and extracted_info["material"]:
        mat_val = extracted_info["material"]
        # 若材料名稱中內含規格數值 (例如 AL6061 160x160x10 或 160*160*10)
        sz_in_mat = extract_blank_size_from_group_name(mat_val)
        if sz_in_mat:
            return sz_in_mat
        # 若使用者直接在材料資料夾寫入單純文字 (如 ExcelTool 的 {material} 填法)
        return mat_val

    # 2. 程式群組名稱
    for chunk in raw_chunks:
        grp = chunk.get("group_name", "")
        sz = extract_blank_size_from_group_name(grp)
        if sz:
            return sz

    # 3. CAM WORKPIECE 素材幾何
    cam_sz = extract_blank_size_from_cam(work_part, uf_session)
    if cam_sz:
        return cam_sz

    # 4. 若無資訊則留空
    return ""

# ==================== 刀具與幾何萃取模組 ====================

def get_tool_from_operation(op, uf_session):
    """
    使用 MachineTool 視圖層級溯源與 UF 備援取得刀具物件與標籤
    """
    tool_obj = None
    tool_tag = 0

    try:
        parent = op.GetParent(NXOpen.CAM.CAMSetup.View.MachineTool)
        while parent is not None:
            if isinstance(parent, NXOpen.CAM.Tool):
                tool_obj = parent
                tool_tag = parent.Tag
                break
            try:
                parent = parent.GetParent()
            except Exception:
                break
    except Exception:
        pass

    if tool_tag == 0:
        try:
            c_tag = uf_session.Oper.AskCutterGroup(op.Tag)
            if c_tag and c_tag != 0:
                tool_tag = c_tag
        except Exception:
            pass

    return tool_obj, tool_tag

def format_tool_number(val):
    """
    標準化刀號格式為 T01, T02...
    """
    if not val or val == "-":
        return "-"
    try:
        num = int(float(val))
        if num > 0:
            return f"T{num:02d}"
    except Exception:
        pass
    return str(val)

def calculate_safe_holder_length(flute_len_str, holder_len_str):
    """
    夾長安全防護運算 (基準至少 20mm，且夾長必須大於刃長 + 5mm)
    """
    try:
        f_val = float(flute_len_str) if flute_len_str != "-" else 0.0
    except Exception:
        f_val = 0.0

    try:
        h_val = float(holder_len_str) if holder_len_str != "-" else 0.0
    except Exception:
        h_val = 0.0

    if h_val > 0:
        if h_val < 20.0:
            h_val = 20.0
        if f_val > 0 and h_val <= f_val:
            h_val = f_val + 5.0
        return f"{h_val:.1f}" if h_val % 1 != 0 else str(int(h_val))
    elif f_val > 0:
        safe_h = max(20.0, f_val + 5.0)
        return f"{safe_h:.1f}" if safe_h % 1 != 0 else str(int(safe_h))

    return holder_len_str

def get_tool_parameters(uf_session, tool_obj, tool_tag):
    """
    透過 UF 底層核心索引讀取刀具參數 (1000直徑, 1001長度, 1002刃長, 1038刀號)
    """
    tool_name = "未指派"
    tool_number = "-"
    tool_diameter = "-"
    flute_length = "-"
    holder_length = "-"

    if tool_obj is not None:
        try:
            tool_name = tool_obj.Name
        except Exception:
            pass
    elif tool_tag and tool_tag != 0:
        try:
            tool_name = uf_session.Obj.AskName(tool_tag)
        except Exception:
            pass

    if not tool_tag or tool_tag == 0:
        return tool_name, tool_number, tool_diameter, flute_length, holder_length

    # 1. 取得刀號
    if tool_obj is not None:
        try:
            if hasattr(tool_obj, "ToolNumber") and tool_obj.ToolNumber > 0:
                tool_number = str(tool_obj.ToolNumber)
        except Exception:
            pass

    if tool_number == "-":
        try:
            t_num = uf_session.Param.AskIntValue(tool_tag, 1038)
            if t_num > 0:
                tool_number = str(t_num)
        except Exception:
            pass

    if tool_number == "-":
        try:
            t_num = uf_session.Obj.AskIntAttr(tool_tag, "TL_NUMBER")
            if t_num > 0:
                tool_number = str(t_num)
        except Exception:
            pass

    if tool_number == "-":
        try:
            t_val = uf_session.Obj.AskAttrValue(tool_tag, "TL_NUMBER")
            if t_val and t_val.strip():
                tool_number = t_val.strip()
        except Exception:
            pass

    tool_number = format_tool_number(tool_number)

    # 2. 透過 Cutter.AskParameters 讀取直徑、長度、刃長
    try:
        params = uf_session.Cutter.AskParameters(tool_tag)
        if params and len(params) > 0:
            if params[0] > 0:
                tool_diameter = f"{params[0]:.2f}"
            if len(params) > 1 and params[1] > 0:
                holder_length = f"{params[1]:.1f}"
            if len(params) > 2 and params[2] > 0:
                flute_length = f"{params[2]:.1f}"
    except Exception:
        pass

    # 3. 備援底層索引讀取
    if tool_diameter == "-":
        try:
            dia = uf_session.Param.AskDoubleValue(tool_tag, 1000)
            if dia > 0:
                tool_diameter = f"{dia:.2f}"
        except Exception:
            pass

    if holder_length == "-":
        try:
            hlen = uf_session.Param.AskDoubleValue(tool_tag, 1001)
            if hlen > 0:
                holder_length = f"{hlen:.1f}"
        except Exception:
            pass

    if flute_length == "-":
        try:
            flen = uf_session.Param.AskDoubleValue(tool_tag, 1002)
            if flen > 0:
                flute_length = f"{flen:.1f}"
        except Exception:
            pass

    holder_length = calculate_safe_holder_length(flute_length, holder_length)

    return tool_name, tool_number, tool_diameter, flute_length, holder_length

def format_tool_display(tool_name, tool_diameter):
    """
    刀具規格顯示格式化：
    若有刀具名稱型號 (如 E3, E1, CC6, C0.5, D0.5) 則直接顯示名稱，不須額外顯示後綴 (Dxx)；
    若無刀具名稱僅有直徑時，顯示為 Dxx (如 D0.5, D0.36)。
    """
    has_name = bool(tool_name and str(tool_name).strip() not in ["未指派", "None", "-", ""])
    has_dia = bool(tool_diameter and str(tool_diameter).strip() not in ["-", "None", ""])

    if has_name:
        clean_name = str(tool_name).strip()
        # 清除任何可能存在的 (Dxx) 或 (xx) 直徑後綴，確保輸出純淨刀名
        clean_name = re.sub(r'\s*\([dD]?[\d\.]+\)$', '', clean_name).strip()
        return clean_name if clean_name else str(tool_name).strip()
    elif has_dia:
        try:
            dia_f = float(tool_diameter)
            dia_str = f"D{dia_f:.2f}".rstrip('0').rstrip('.') if dia_f % 1 != 0 else f"D{int(dia_f)}"
            return dia_str
        except Exception:
            return f"D{tool_diameter}"
    return "-"

def get_operation_time_seconds(op):
    """
    取得工序切削時間 (單位：秒)
    """
    minutes = 0.0
    try:
        minutes = op.GetToolpathTime()
    except Exception:
        pass

    if minutes <= 0.0:
        try:
            minutes = op.GetToolpathCuttingTime()
        except Exception:
            pass

    total_seconds = int(round(minutes * 60.0))
    return max(0, total_seconds)

def format_seconds_to_hms(total_seconds):
    """
    將總秒數格式化為 HH:MM:SS
    """
    if total_seconds <= 0:
        return "00:00:00"
    hours, remainder = divmod(int(total_seconds), 3600)
    mins, secs = divmod(remainder, 60)
    return f"{hours:02d}:{mins:02d}:{secs:02d}"

def get_operation_time(op):
    """
    取得工序切削時間並格式化為 HH:MM:SS
    """
    return format_seconds_to_hms(get_operation_time_seconds(op))

# ==================== 刀號與工步智慧合併模組 (Tool Merge Engine) ====================

def parse_tool_int(val):
    """
    解析刀號字串或數字為整數，若無效則回傳 -99999
    """
    if not val or str(val).strip() in ["-", "", "None"]:
        return -99999
    try:
        m = re.search(r'\d+', str(val))
        if m:
            return int(m.group())
    except Exception:
        pass
    return -99999

def normalize_tool_name(name):
    """
    標準化刀具名稱：移除尾部 -Txx 或 _Txx (例如 E4-T01 -> E4)
    """
    if not name or name in ["未指派", "None", "-"]:
        return ""
    clean = str(name).strip()
    clean = re.sub(r'[-_]T\d+$', '', clean, flags=re.IGNORECASE)
    return clean

def extract_float(val):
    """
    安全解析浮點數數值
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

def merge_consecutive_tools(raw_ops, group_name=""):
    """
    依據 ExcelTool 核心規則合併同群組相鄰刀具工步：
    1. 狀況 A (同刀號累加)：相鄰工步刀號相同時，合併為一列，加工時間自動加總
    2. 狀況 B (連續刀號同規格)：相鄰工步刀號連續 (t_curr == t_prev + 1)，
       且刀名去後綴、直徑、刃長、夾長一致時，合併為範圍字串 (如 T01~T02)，加工時間加總，長度取最大值防護
    3. 備註欄位串接合併，保留轉速進給資訊
    4. 特殊防護 (單獨點選工序)：若前一工步或當前工步為「直接單獨選取之工序」(is_direct_op == True)，
       代表工程師指定獨立轉出與呈現該工序，絕不參與同刀號/連續刀具合併，每道工序各自獨立成列並顯示其工序名稱！
    """
    if not raw_ops:
        return []

    merged_ops = []
    temp_op = None

    for op in raw_ops:
        if temp_op is None:
            temp_op = dict(op)
            t_int = parse_tool_int(temp_op.get("tool_number"))
            temp_op["_start_tool_str"] = format_tool_number(temp_op.get("tool_number"))
            temp_op["_last_tool_int"] = t_int
            continue

        # 0. 若前一個工步或當前工步為使用者「直接點選之單獨工序」，禁止合併，各自獨立成列
        is_direct_curr = op.get("is_direct_op", False)
        is_direct_prev = temp_op.get("is_direct_op", False)
        if is_direct_curr or is_direct_prev:
            merged_ops.append(temp_op)
            temp_op = dict(op)
            t_int = parse_tool_int(temp_op.get("tool_number"))
            temp_op["_start_tool_str"] = format_tool_number(temp_op.get("tool_number"))
            temp_op["_last_tool_int"] = t_int
            continue

        # 1. 取得前後刀號整數
        t_curr_int = parse_tool_int(op.get("tool_number"))
        t_prev_int = temp_op.get("_last_tool_int", parse_tool_int(temp_op.get("tool_number")))

        # 2. 規格一致性比對 (乾淨刀名、直徑、刃長、夾長)
        name_prev = normalize_tool_name(temp_op.get("raw_tool_name", temp_op.get("tool_name", "")))
        name_curr = normalize_tool_name(op.get("raw_tool_name", op.get("tool_name", "")))

        dia_prev = extract_float(temp_op.get("tool_diameter"))
        dia_curr = extract_float(op.get("tool_diameter"))

        flute_prev = extract_float(temp_op.get("flute_length"))
        flute_curr = extract_float(op.get("flute_length"))

        holder_prev = extract_float(temp_op.get("holder_length"))
        holder_curr = extract_float(op.get("holder_length"))

        is_spec_match = (
            name_prev == name_curr and
            dia_prev == dia_curr and
            flute_prev == flute_curr and
            holder_prev == holder_curr
        )

        # 狀況 A：同把刀 (刀號相同且有效)
        is_same_id = (t_curr_int == t_prev_int and t_curr_int != -99999)

        # 狀況 B：連續刀號且規格完全相同 (例如 T01 與 T02)
        is_consecutive = (t_curr_int == t_prev_int + 1 and t_curr_int != -99999 and is_spec_match)

        if is_same_id or is_consecutive:
            # 觸發合併
            # (1) 刀名正規化為乾淨名稱 (去後綴)
            if name_prev:
                temp_op["raw_tool_name"] = name_prev
                temp_op["tool_name"] = name_prev

            # (2) 刀號字串處理
            if is_consecutive:
                start_str = temp_op.get("_start_tool_str", format_tool_number(temp_op.get("tool_number")))
                end_str = format_tool_number(op.get("tool_number"))
                temp_op["tool_number"] = f"{start_str}~{end_str}"
                temp_op["_last_tool_int"] = t_curr_int
            else:
                # 相同刀號維持原樣
                temp_op["_last_tool_int"] = t_curr_int

            # (3) 加工時間秒數累加
            sec_prev = temp_op.get("time_seconds", 0)
            sec_curr = op.get("time_seconds", 0)
            total_sec = sec_prev + sec_curr
            temp_op["time_seconds"] = total_sec
            temp_op["time"] = format_seconds_to_hms(total_sec)

            # (4) 刃長與夾長取最大值防護
            max_flute = max(flute_prev, flute_curr)
            max_holder = max(holder_prev, holder_curr)
            flute_str = f"{max_flute:.1f}" if max_flute % 1 != 0 else str(int(max_flute)) if max_flute > 0 else "-"
            holder_str = f"{max_holder:.1f}" if max_holder % 1 != 0 else str(int(max_holder)) if max_holder > 0 else "-"
            temp_op["flute_length"] = flute_str
            temp_op["holder_length"] = calculate_safe_holder_length(flute_str, holder_str)

            # (5) 重新格式化刀具規格欄位
            clean_display_name = temp_op.get("raw_tool_name", temp_op.get("tool_name", ""))
            temp_op["tool_spec_display"] = format_tool_display(clean_display_name, temp_op.get("tool_diameter"))

            # (6) 備註串接合併
            n1 = str(temp_op.get("note", "")).strip()
            n2 = str(op.get("note", "")).strip()
            if n2 and n2 not in n1:
                temp_op["note"] = f"{n1} {n2}".strip()
        else:
            # 不符合合併條件，結算前一工步
            merged_ops.append(temp_op)
            temp_op = dict(op)
            t_int = parse_tool_int(temp_op.get("tool_number"))
            temp_op["_start_tool_str"] = format_tool_number(temp_op.get("tool_number"))
            temp_op["_last_tool_int"] = t_int

    if temp_op is not None:
        merged_ops.append(temp_op)

    # 重新編排群組工站序號 (seq)、程式檔名填寫規則 (直接選取顯示工序名，資料夾展開首行填寫後續留空) 與規格統一格式化
    final_ops = []
    for idx, mop in enumerate(merged_ops):
        seq_num = idx + 1
        is_direct = mop.get("is_direct_op", False)
        if is_direct:
            prog_file_name = mop.get("clean_op_name", "")
        else:
            if idx == 0:
                prog_file_name = clean_program_name(group_name) if (group_name and group_name != "DEFAULT") else mop.get("clean_op_name", "")
            else:
                prog_file_name = ""

        mop["seq"] = seq_num
        mop["op_name"] = prog_file_name

        # 統一格式化刀具規格欄位 (如 E4 (D4.0))
        clean_tool_title = mop.get("raw_tool_name", mop.get("tool_name", ""))
        mop["tool_spec_display"] = format_tool_display(clean_tool_title, mop.get("tool_diameter"))

        final_ops.append(mop)

    return final_ops

# ==================== 工段示圖截取與拍照精靈模組 ====================

def get_distinct_stages(chunks):
    """
    分析所有 chunks 歸納出獨立工段清單
    若有明顯工段標籤 (如 M1, M2, OP10) 則保留各工段；
    若無特定標籤，則歸納為單一工段，避免多餘拍照彈窗干擾。
    """
    raw_stages = []
    for c in chunks:
        if not c.get("operations"):
            continue
        stg = c.get("stage")
        if not stg:
            stg = extract_stage_key(c.get("group_name", "")) or "通用工段"
        if stg not in raw_stages:
            raw_stages.append(stg)

    if not raw_stages:
        return ["通用工段"]

    # 檢查是否有具體的工段識別標籤 (M數字, OP數字 等)
    has_explicit_stage = any(re.match(r'^(M\d+|OP\d+)$', k) for k in raw_stages)
    if has_explicit_stage:
        explicit_list = [s for s in raw_stages if s != "通用工段"]
        return explicit_list if explicit_list else ["通用工段"]

    return ["通用工段"]

# ==================== 跨機器環境與組件智慧解析引擎 ====================
_CACHED_PYTHON_RUNTIME_GUI = None
_CACHED_PYTHON_RUNTIME_NOGUI = None

def resolve_project_root_dir(the_session=None, work_part=None):
    """
    智慧解析專案根目錄 (相容 NX Journal 模式、任意磁碟/路徑佈署、隨身碟或工作站環境)
    優先順序：
    1. globals().__file__ (若以獨立腳本或 IDE 執行)
    2. the_session.ExecutingJournal (NX Journal 模式官方標準屬性)
    3. sys.argv[0] (部分 NX 版本傳入的 journal 完整路徑)
    4. os.getcwd() (當前工作目錄)
    5. work_part.FullPath 所在目錄及其 Template 子目錄
    6. 標準備援路徑 C:\\NX_Standard\\Template, D:\\NX_Standard\\Template
    """
    candidate_dirs = []

    # 1. 檢查 __file__
    if "__file__" in globals() and globals()["__file__"]:
        try:
            candidate_dirs.append(os.path.dirname(os.path.abspath(globals()["__file__"])))
        except Exception:
            pass

    # 2. 檢查 NX Session 的 ExecutingJournal 屬性
    if the_session is not None:
        try:
            exec_j = getattr(the_session, "ExecutingJournal", None)
            if exec_j and os.path.exists(exec_j):
                candidate_dirs.append(os.path.dirname(os.path.abspath(exec_j)))
        except Exception:
            pass

    # 3. 檢查 sys.argv[0]
    if sys.argv and sys.argv[0]:
        try:
            cand = os.path.abspath(sys.argv[0])
            if os.path.exists(cand):
                candidate_dirs.append(os.path.dirname(cand) if os.path.isfile(cand) else cand)
        except Exception:
            pass

    # 4. 檢查當前工作目錄
    try:
        candidate_dirs.append(os.getcwd())
    except Exception:
        pass

    # 5. 檢查當前工作 Part 所在目錄及其同級/上級目錄
    if work_part is not None:
        try:
            part_path = getattr(work_part, "FullPath", None)
            if part_path and os.path.exists(part_path):
                part_dir = os.path.dirname(os.path.abspath(part_path))
                candidate_dirs.append(part_dir)
                candidate_dirs.append(os.path.join(part_dir, "Template"))
        except Exception:
            pass

    # 6. 標準固定路徑
    candidate_dirs.append(r"C:\NX_Standard\Template")
    candidate_dirs.append(r"D:\NX_Standard\Template")

    # 去重並驗證目錄中是否存在核心組件或範本
    seen = set()
    best_dir = None
    key_files = ["ShopDoc_Template.xlsx", "nc_post_dialog.py", "nx_cam_to_excel.py"]

    for d in candidate_dirs:
        if not d:
            continue
        try:
            d_norm = os.path.normpath(os.path.abspath(d))
            if d_norm.lower() in seen:
                continue
            seen.add(d_norm.lower())

            if os.path.isdir(d_norm):
                hits = sum(1 for kf in key_files if os.path.exists(os.path.join(d_norm, kf)))
                if hits > 0:
                    best_dir = d_norm
                    break
        except Exception:
            continue

    if not best_dir:
        best_dir = r"C:\NX_Standard\Template"

    return best_dir

def resolve_asset_file(filename, the_session=None, work_part=None):
    """
    精確解析專案資產檔案 (對話框腳本、Excel 範本等) 之絕對路徑
    """
    root_dir = resolve_project_root_dir(the_session=the_session, work_part=work_part)
    target_path = os.path.join(root_dir, filename)
    if os.path.exists(target_path):
        return target_path

    # 若根目錄未直接命中，進一步搜尋常見位置
    search_dirs = [
        os.getcwd(),
        r"C:\NX_Standard\Template",
        r"D:\NX_Standard\Template"
    ]
    if "__file__" in globals() and globals()["__file__"]:
        try:
            search_dirs.insert(0, os.path.dirname(os.path.abspath(globals()["__file__"])))
        except Exception:
            pass

    for s_dir in search_dirs:
        try:
            p = os.path.join(s_dir, filename)
            if os.path.exists(p):
                return p
        except Exception:
            continue

    return target_path

def resolve_python_runtime(require_gui=True):
    """
    全方位跨機器 Python 直譯器智慧探測與純淨環境建構器
    支援：
    1. Windows 註冊表掃描 (Python 3.8 ~ 3.14+)
    2. 常見安裝目錄列舉 (AppData, ProgramFiles, Anaconda, Miniconda)
    3. Windows Python Launcher (pyw.exe / py.exe -3)
    4. 系統 PATH 探測 (安全排除 WindowsApps 微軟商店假捷徑)
    5. 當前直譯器 sys.executable
    6. 針對 GUI 對話框進行輕量級 Tkinter 可用性快速校驗
    7. 徹底隔離 NX 專屬的 PYTHONHOME / PYTHONPATH 與 Tcl/Tk 版本衝突
    回傳: (interpreter_cmd_list, clean_env)
    """
    global _CACHED_PYTHON_RUNTIME_GUI, _CACHED_PYTHON_RUNTIME_NOGUI

    if require_gui and _CACHED_PYTHON_RUNTIME_GUI is not None:
        cmd, env = _CACHED_PYTHON_RUNTIME_GUI
        return list(cmd), env.copy()
    if not require_gui and _CACHED_PYTHON_RUNTIME_NOGUI is not None:
        cmd, env = _CACHED_PYTHON_RUNTIME_NOGUI
        return list(cmd), env.copy()

    candidates = []

    # 1. 掃描 Windows 註冊表 (HKCU 與 HKLM)
    if winreg:
        for root in [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]:
            for key_path in [r"Software\Python\PythonCore", r"Software\WOW6432Node\Python\PythonCore"]:
                try:
                    with winreg.OpenKey(root, key_path) as base_key:
                        num_subkeys = winreg.QueryInfoKey(base_key)[0]
                        for i in range(num_subkeys):
                            try:
                                ver_name = winreg.EnumKey(base_key, i)
                                if ver_name.startswith("2."):
                                    continue
                                with winreg.OpenKey(base_key, rf"{ver_name}\InstallPath") as ip_key:
                                    install_dir, _ = winreg.QueryValueEx(ip_key, "")
                                    if install_dir and os.path.isdir(install_dir):
                                        pw = os.path.join(install_dir, "pythonw.exe")
                                        p = os.path.join(install_dir, "python.exe")
                                        if os.path.exists(pw):
                                            candidates.append([pw])
                                        if os.path.exists(p):
                                            candidates.append([p])
                            except Exception:
                                pass
                except Exception:
                    pass

    # 2. 智慧列舉常見安裝目錄
    patterns = [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Python\Python3*"),
        os.path.join(os.environ.get("ProgramFiles", ""), r"Python3*"),
        os.path.join(os.environ.get("ProgramFiles(x86)", ""), r"Python3*"),
        r"C:\Python3*",
        r"D:\Python3*",
        os.path.join(os.environ.get("ProgramData", ""), r"anaconda*"),
        os.path.join(os.environ.get("ProgramData", ""), r"miniconda*"),
        os.path.join(os.environ.get("USERPROFILE", ""), r"anaconda*"),
        os.path.join(os.environ.get("USERPROFILE", ""), r"miniconda*")
    ]
    for pat in patterns:
        try:
            for p_dir in glob.glob(pat):
                if os.path.isdir(p_dir):
                    pw = os.path.join(p_dir, "pythonw.exe")
                    p = os.path.join(p_dir, "python.exe")
                    if os.path.exists(pw):
                        candidates.append([pw])
                    if os.path.exists(p):
                        candidates.append([p])
        except Exception:
            pass

    # 3. 掃描 Windows Python Launcher (pyw.exe / py.exe)
    py_launcher = shutil.which("pyw") or shutil.which("py")
    if not py_launcher:
        for p_cand in [
            r"C:\Windows\pyw.exe",
            r"C:\Windows\py.exe",
            os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Python\Launcher\pyw.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Python\Launcher\py.exe")
        ]:
            try:
                if os.path.exists(p_cand):
                    py_launcher = p_cand
                    break
            except Exception:
                pass
    if py_launcher:
        candidates.append([py_launcher, "-3"])

    # 4. 掃描 PATH (排除 WindowsApps 假捷徑)
    for cmd in ["pythonw", "python"]:
        try:
            found = shutil.which(cmd)
            if found and "windowsapps" not in found.lower():
                candidates.append([found])
        except Exception:
            pass

    # 5. 當前直譯器 (若是 NX 內建 Python 或系統 Python)
    if sys.executable and os.path.exists(sys.executable):
        candidates.append([sys.executable])

    seen_keys = set()
    CREATE_NO_WINDOW = 0x08000000

    for cand_cmd in candidates:
        cand_key = tuple(cand_cmd)
        if cand_key in seen_keys:
            continue
        seen_keys.add(cand_key)

        clean_env = os.environ.copy()
        clean_env.pop("PYTHONHOME", None)
        clean_env.pop("PYTHONPATH", None)
        clean_env.pop("PYTHONSTARTUP", None)
        clean_env.pop("PYTHONEXECUTABLE", None)

        exe_path = cand_cmd[0]
        if os.path.isabs(exe_path) and os.path.exists(exe_path):
            py_dir = os.path.dirname(exe_path)
            tcl_lib = os.path.join(py_dir, r"tcl\tcl8.6")
            tk_lib = os.path.join(py_dir, r"tcl\tk8.6")
            if os.path.exists(tcl_lib):
                clean_env["TCL_LIBRARY"] = tcl_lib
            else:
                clean_env.pop("TCL_LIBRARY", None)
            if os.path.exists(tk_lib):
                clean_env["TK_LIBRARY"] = tk_lib
            else:
                clean_env.pop("TK_LIBRARY", None)

            py_dlls = os.path.join(py_dir, "DLLs")
            py_scripts = os.path.join(py_dir, "Scripts")
            filtered_paths = [p for p in clean_env.get("PATH", "").split(";") if p and "nxbin\\python" not in p.lower()]
            clean_env["PATH"] = ";".join([py_dlls, py_dir, py_scripts] + filtered_paths)

        if require_gui:
            test_cmd = list(cand_cmd) + ["-c", "import tkinter"]
            try:
                res = subprocess.run(
                    test_cmd,
                    capture_output=True,
                    timeout=2,
                    env=clean_env,
                    creationflags=CREATE_NO_WINDOW
                )
                if res.returncode == 0:
                    _CACHED_PYTHON_RUNTIME_GUI = (cand_cmd, clean_env)
                    return list(cand_cmd), clean_env
            except Exception:
                continue
        else:
            _CACHED_PYTHON_RUNTIME_NOGUI = (cand_cmd, clean_env)
            return list(cand_cmd), clean_env

    # 若未找到任何可用直譯器，真實回傳 None (由 NX 原生對話框接手處理)
    fallback_env = os.environ.copy()
    fallback_env.pop("PYTHONHOME", None)
    fallback_env.pop("PYTHONPATH", None)
    if require_gui:
        _CACHED_PYTHON_RUNTIME_GUI = (None, fallback_env)
    else:
        _CACHED_PYTHON_RUNTIME_NOGUI = (None, fallback_env)
    return None, fallback_env

def get_clean_subprocess_env():
    """
    維持現有調用介面相容性的包裝函式
    回傳: (valid_pythonw_str_or_list, clean_env)
    """
    cmd, env = resolve_python_runtime(require_gui=True)
    if cmd:
        valid_pythonw = cmd[0] if len(cmd) == 1 else cmd
        return valid_pythonw, env
    return None, env

def ensure_white_background_image(png_path):
    """
    確保圖檔背景為 100% 純白色 (RGB 255, 255, 255)：
    1. 若圖檔包含 Alpha 透明通道，複合至純白底色；
    2. 若圖檔為 RGB 但四周邊界非純白色 (如 NX 漸變背景未清除)，執行邊界多點泛洪填充 (Floodfill) 將外圍背景清洗為純白。
    雙層防護：支援當前直譯器 PIL 與外部 Python (已安裝 Pillow) 執行。
    """
    if not os.path.exists(png_path) or os.path.getsize(png_path) == 0:
        return

    # 優先嘗試在當前直譯器中運行
    try:
        from PIL import Image, ImageDraw
        with Image.open(png_path) as im:
            im_rgba = im.convert("RGBA")
            bg = Image.new("RGBA", im_rgba.size, (255, 255, 255, 255))
            bg.paste(im_rgba, mask=im_rgba.split()[-1])
            rgb_im = bg.convert("RGB")
            w, h = rgb_im.size
            corners = [rgb_im.getpixel((0, 0)), rgb_im.getpixel((w-1, 0)), rgb_im.getpixel((0, h-1)), rgb_im.getpixel((w-1, h-1))]
            if any(sum(c) < 735 for c in corners):
                step = max(10, min(w, h) // 40)
                seeds = (
                    [(x, 0) for x in range(0, w, step)] +
                    [(x, h-1) for x in range(0, w, step)] +
                    [(0, y) for y in range(0, h, step)] +
                    [(w-1, y) for y in range(0, h, step)]
                )
                for pt in seeds:
                    if sum(rgb_im.getpixel(pt)) < 735:
                        try:
                            ImageDraw.floodfill(rgb_im, pt, (255, 255, 255), thresh=45)
                        except Exception:
                            pass
            rgb_im.save(png_path, "PNG")
            try:
                rgb_im.close()
                bg.close()
                im_rgba.close()
            except Exception:
                pass
            return
    except Exception:
        pass

    # 若當前直譯器未安裝 PIL，呼叫外部 Python 執行 (延遲構建指令碼)
    try:
        external_script = (
            "from PIL import Image, ImageDraw\n"
            "def process(path):\n"
            "    with Image.open(path) as im:\n"
            "        im = im.convert('RGBA')\n"
            "        bg = Image.new('RGBA', im.size, (255, 255, 255, 255))\n"
            "        bg.paste(im, mask=im.split()[-1])\n"
            "        rgb_im = bg.convert('RGB')\n"
            "        w, h = rgb_im.size\n"
            "        corners = [rgb_im.getpixel((0, 0)), rgb_im.getpixel((w-1, 0)), rgb_im.getpixel((0, h-1)), rgb_im.getpixel((w-1, h-1))]\n"
            "        need_flood = any(sum(c) < 735 for c in corners)\n"
            "        if need_flood:\n"
            "            step = max(10, min(w, h) // 40)\n"
            "            seeds = ([(x, 0) for x in range(0, w, step)] +\n"
            "                     [(x, h-1) for x in range(0, w, step)] +\n"
            "                     [(0, y) for y in range(0, h, step)] +\n"
            "                     [(w-1, y) for y in range(0, h, step)])\n"
            "            for pt in seeds:\n"
            "                if sum(rgb_im.getpixel(pt)) < 735:\n"
            "                    try:\n"
            "                        ImageDraw.floodfill(rgb_im, pt, (255, 255, 255), thresh=45)\n"
            "                    except Exception:\n"
            "                        pass\n"
            "        rgb_im.save(path, 'PNG')\n"
            f"process(r'{png_path}')\n"
        )
        py_cmd, py_env = resolve_python_runtime(require_gui=False)
        if py_cmd:
            full_cmd = list(py_cmd) + ["-c", external_script]
            subprocess.run(full_cmd, timeout=6, env=py_env, creationflags=0x08000000)
    except Exception:
        pass

def capture_nx_viewport(the_ui, out_png_path, white_background=True, listing=None, work_part=None):
    """
    截取當前 NX 圖形視窗畫面為高解析度 PNG 圖片 (預設白底渲染模式，避免黑底浪費墨水)
    技術特色：
    1. 底層座標圖示徹底隱藏：
       - 視圖層級 (WorkView.TriadVisibility & WcsVisibility)：直接關閉畫面左下角 3D 方塊方位座標與工作座標系
       - 零件偏好設定層級 (PartPreferences.ScreenVisualization.TriadVisibility)：關閉當前零件視圖三面角
       - 會話偏好設定層級 (SessionPreferences.ScreenVisualization.TriadVisibility)：關閉 Session 視圖三面角
       - 座標系控制器 (WCS.Visibility & UFSession.Csys.SetWcsDisplay)：雙重隱藏 WCS
       - 即時視圖刷新 (WorkView.Update & UFSession.Disp.RegenerateDisplay)：立即清除渲染緩衝區中的座標殘留
    2. 拍照前透過 NX 原生 CreateBackground 暫時切換視圖為單色純白底，截圖後自動還原
    3. 配合 ImageExportBuilder Transparent 模式輸出
    4. 圖片產生後調用 ensure_white_background_image 進行邊界泛洪去背，確保 100% 絕對純白底！
    5. 全程透過 try...finally 架構，拍照結束後 100% 精確還原使用者原本的視圖座標與背景設定！
    """
    out_dir = os.path.dirname(out_png_path)
    if out_dir and not os.path.exists(out_dir):
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception:
            pass

    base_no_ext, ext = os.path.splitext(out_png_path)
    if not ext:
        ext = ".png"
        out_png_path = base_no_ext + ext

    # 安全動態匯入 Gateway 模組
    NXGateway = None
    try:
        import NXOpen.Gateway as _gw
        NXGateway = _gw
    except Exception:
        try:
            import NXOpen_Gateway as _gw
            NXGateway = _gw
        except Exception:
            NXGateway = None

    # 確保取得 Session、WorkPart、WorkView 與 UFSession
    the_session = None
    try:
        the_session = NXOpen.Session.GetSession()
        if work_part is None and hasattr(the_session, "Parts") and hasattr(the_session.Parts, "Work"):
            work_part = the_session.Parts.Work
    except Exception:
        pass

    w_view = None
    if work_part is not None and hasattr(work_part, "Views"):
        try:
            w_view = work_part.Views.WorkView
        except Exception:
            w_view = None

    uf_session = None
    try:
        uf_session = NXOpen.UF.UFSession.GetUFSession()
    except Exception:
        pass

    # 暫時隱藏 View Triad (視圖方塊三軸座標架) 與 WCS (工作座標系)
    orig_view_triad_vis = None
    orig_view_wcs_vis = None
    orig_part_triad_vis = None
    orig_sess_triad_vis = None
    orig_wcs_vis = None
    part_screen = None
    sess_screen = None
    bg_override = None
    orig_bg_type = None

    try:
        # (A) 視圖層級控制 (View Level：最底層直接控制當前視窗 Triad 與 WCS 渲染)
        if w_view is not None:
            try:
                if hasattr(w_view, "TriadVisibility"):
                    orig_view_triad_vis = w_view.TriadVisibility
                    w_view.TriadVisibility = False
            except Exception:
                pass
            try:
                if hasattr(w_view, "WcsVisibility"):
                    orig_view_wcs_vis = w_view.WcsVisibility
                    w_view.WcsVisibility = False
            except Exception:
                pass

        # (B) 零件偏好設定層級 (Part Preferences Level：NX 12+ 官方標準)
        if work_part is not None and hasattr(work_part, "Preferences"):
            try:
                if hasattr(work_part.Preferences, "ScreenVisualization"):
                    part_screen = work_part.Preferences.ScreenVisualization
                    if hasattr(part_screen, "TriadVisibility"):
                        orig_part_triad_vis = part_screen.TriadVisibility
                        part_screen.TriadVisibility = False
            except Exception:
                pass

        # (C) 會話偏好設定層級 (Session Preferences Level：全域備援)
        if the_session is not None and hasattr(the_session, "Preferences"):
            try:
                if hasattr(the_session.Preferences, "ScreenVisualization"):
                    sess_screen = the_session.Preferences.ScreenVisualization
                    if hasattr(sess_screen, "TriadVisibility"):
                        orig_sess_triad_vis = sess_screen.TriadVisibility
                        sess_screen.TriadVisibility = 0
            except Exception:
                pass

        # (D) 工作座標系 (WCS / UF Csys 控制)
        if work_part is not None and hasattr(work_part, "WCS"):
            try:
                if hasattr(work_part.WCS, "Visibility"):
                    orig_wcs_vis = work_part.WCS.Visibility
                    work_part.WCS.Visibility = False
            except Exception:
                pass
        if uf_session is not None and hasattr(uf_session, "Csys"):
            try:
                if hasattr(uf_session.Csys, "SetWcsDisplay"):
                    uf_session.Csys.SetWcsDisplay(0)
            except Exception:
                pass

        # (E) 強制立即重繪與刷新視圖，使隱藏 Triad / WCS 立即生效於渲染緩衝區
        try:
            if w_view is not None and hasattr(w_view, "Update"):
                w_view.Update()
        except Exception:
            pass
        try:
            if uf_session is not None and hasattr(uf_session, "Disp"):
                if hasattr(uf_session.Disp, "RegenerateDisplay"):
                    uf_session.Disp.RegenerateDisplay()
                if hasattr(uf_session.Disp, "Refresh"):
                    uf_session.Disp.Refresh()
        except Exception:
            pass

        # 第一重防護：截圖前透過 NX 原生 API 將視圖暫時設為純白底
        if white_background and work_part is not None and w_view is not None:
            try:
                bg_override = work_part.Views.CreateBackground(w_view, False)
                try:
                    orig_bg_type = bg_override.BackgroundShadedViewsType
                except Exception:
                    pass
                bg_override.BackgroundShadedViewsType = 1 # 1: Plain (單色背景)
                bg_override.SetBackgroundShadedViewsPlain([1.0, 1.0, 1.0]) # 純白色
                bg_override.Commit()
            except Exception:
                bg_override = None

        # 1. 現代 NX 標準 API: CreateImageExportBuilder
        builder = None
        try:
            if the_ui is not None and hasattr(the_ui, "CreateImageExportBuilder"):
                builder = the_ui.CreateImageExportBuilder()
        except Exception:
            pass

        if builder is None and work_part is not None:
            try:
                if hasattr(work_part, "Views") and hasattr(work_part.Views, "CreateImageExportBuilder"):
                    builder = work_part.Views.CreateImageExportBuilder()
            except Exception:
                pass

        if builder is not None:
            try:
                builder.RegionMode = False

                # 設定輸出格式為 PNG
                format_set = False
                if NXGateway and hasattr(NXGateway, "ImageExportBuilder"):
                    try:
                        builder.FileFormat = NXGateway.ImageExportBuilder.FileFormats.Png
                        format_set = True
                    except Exception:
                        pass
                if not format_set:
                    try:
                        builder.FileFormat = NXOpen.Gateway.ImageExportBuilder.FileFormats.Png
                        format_set = True
                    except Exception:
                        pass

                # 設定目標檔名
                builder.FileName = out_png_path

                # 白底輸出設定：優先使用 Transparent (透明背景去背，再複合為純白底)
                if white_background:
                    bg_set = False
                    if NXGateway and hasattr(NXGateway, "ImageExportBuilder"):
                        try:
                            builder.BackgroundOption = NXGateway.ImageExportBuilder.BackgroundOptions.Transparent
                            bg_set = True
                        except Exception:
                            pass
                    if not bg_set:
                        try:
                            builder.BackgroundOption = NXOpen.Gateway.ImageExportBuilder.BackgroundOptions.Transparent
                            bg_set = True
                        except Exception:
                            pass
                    if not bg_set and NXGateway and hasattr(NXGateway, "ImageExportBuilder"):
                        try:
                            builder.BackgroundOption = NXGateway.ImageExportBuilder.BackgroundOptions.Original
                            bg_set = True
                        except Exception:
                            pass

                # 執行輸出
                builder.Commit()

            except Exception as ex_bld:
                if listing:
                    listing.WriteLine(f"  [ImageExportBuilder 異常] {str(ex_bld)}")
            finally:
                try:
                    builder.Destroy()
                except Exception:
                    pass

        # 檢查是否有檔案生成 (涵蓋 NX 自行追加 .png 或雙重副檔名狀況)
        possible_paths = [
            out_png_path,
            out_png_path + ".png",
            base_no_ext + ".png"
        ]
        has_file = False
        for p in possible_paths:
            if os.path.exists(p) and os.path.getsize(p) > 0:
                if p != out_png_path:
                    try:
                        if os.path.exists(out_png_path):
                            os.remove(out_png_path)
                        os.rename(p, out_png_path)
                    except Exception:
                        out_png_path = p
                has_file = True
                break

        # 2. 備援方案：若 builder 未能輸出圖檔，嘗試 UF.Disp (若支援)
        if not has_file and uf_session is not None:
            try:
                if hasattr(uf_session, "Disp") and hasattr(uf_session.Disp, "CreateImage"):
                    uf_session.Disp.CreateImage(out_png_path, 2, 1) # 2: PNG, 1: White
            except Exception:
                pass
            for p in possible_paths:
                if os.path.exists(p) and os.path.getsize(p) > 0:
                    if p != out_png_path:
                        try:
                            if os.path.exists(out_png_path):
                                os.remove(out_png_path)
                            os.rename(p, out_png_path)
                        except Exception:
                            out_png_path = p
                    has_file = True
                    break

        if has_file:
            # 第二重防護：透明背景複合 + 邊界多點泛洪清洗
            if white_background:
                ensure_white_background_image(out_png_path)

            if listing:
                listing.WriteLine(f"  [示圖擷取成功] 圖檔已儲存 (已強化純白底與座標圖示隱藏)：{out_png_path}")
            return True

        if listing:
            listing.WriteLine(f"  [示圖擷取失敗] 無法在硬碟產生圖檔：{out_png_path}")
        return False

    finally:
        # 還原 NX 原生視圖背景
        if bg_override is not None:
            try:
                if orig_bg_type is not None:
                    bg_override.BackgroundShadedViewsType = orig_bg_type
                    bg_override.Commit()
                bg_override.Destroy()
            except Exception:
                pass

        # 全方位恢復 View Triad (方位方塊座標) 與 WCS 顯示，尊重使用者原本的偏好設定
        restore_triad_and_wcs(
            the_session, work_part, w_view, uf_session,
            orig_view_triad=orig_view_triad_vis,
            orig_view_wcs=orig_view_wcs_vis,
            orig_part_triad=orig_part_triad_vis,
            orig_sess_triad=orig_sess_triad_vis,
            orig_wcs=orig_wcs_vis
        )

def restore_triad_and_wcs(the_session=None, work_part=None, w_view=None, uf_session=None,
                          orig_view_triad=None, orig_view_wcs=None,
                          orig_part_triad=None, orig_sess_triad=None, orig_wcs=None):
    """
    底層座標圖示與視圖方位方塊 (View Triad / WCS) 全方位還原核心：
    優先依據拍照前記錄的原始偏好設定還原，若無記錄 (None) 則安全預設為顯示 (True / 1)。
    1. 視圖層級 (WorkView.TriadVisibility & WcsVisibility)
    2. 零件偏好設定 (PartPreferences.ScreenVisualization)
    3. 會話偏好設定 (SessionPreferences.ScreenVisualization)
    4. 工作座標系 (WCS.Visibility & UFSession.Csys.SetWcsDisplay)
    5. 立即調用 RegenerateDisplay / Refresh / MakeDisplayUpToDate 強制刷新圖形渲染緩衝區！
    """
    if the_session is None:
        try:
            the_session = NXOpen.Session.GetSession()
        except Exception:
            pass

    if work_part is None and the_session is not None and hasattr(the_session, "Parts"):
        try:
            work_part = the_session.Parts.Work
        except Exception:
            pass

    if w_view is None and work_part is not None and hasattr(work_part, "Views"):
        try:
            w_view = work_part.Views.WorkView
        except Exception:
            pass

    if uf_session is None:
        try:
            uf_session = NXOpen.UF.UFSession.GetUFSession()
        except Exception:
            pass

    # (A) 視圖層級 View Triad 與 WCS 恢復 (尊重原始偏好)
    target_view_triad = orig_view_triad if orig_view_triad is not None else True
    target_view_wcs = orig_view_wcs if orig_view_wcs is not None else True

    if w_view is not None:
        try:
            if hasattr(w_view, "TriadVisibility"):
                w_view.TriadVisibility = bool(target_view_triad)
        except Exception:
            pass
        try:
            if hasattr(w_view, "WcsVisibility"):
                w_view.WcsVisibility = bool(target_view_wcs)
        except Exception:
            pass

    # (B) 零件偏好設定層級 (PartPreferences)
    if work_part is not None and hasattr(work_part, "Preferences"):
        pref_objs = []
        for attr in ["ScreenVisualization", "PartVisualizationScreen", "VisualizationScreen"]:
            if hasattr(work_part.Preferences, attr):
                try:
                    pref_objs.append(getattr(work_part.Preferences, attr))
                except Exception:
                    pass
        target_part_triad = orig_part_triad if orig_part_triad is not None else True
        for p_obj in pref_objs:
            try:
                if hasattr(p_obj, "SetTriadVisibility"):
                    p_obj.SetTriadVisibility(bool(target_part_triad))
            except Exception:
                pass
            try:
                if hasattr(p_obj, "TriadVisibility"):
                    p_obj.TriadVisibility = bool(target_part_triad)
            except Exception:
                pass

    # (C) 會話偏好設定層級 (SessionPreferences)
    if the_session is not None and hasattr(the_session, "Preferences"):
        sess_objs = []
        for attr in ["ScreenVisualization", "SessionVisualizationScreen", "VisualizationScreen"]:
            if hasattr(the_session.Preferences, attr):
                try:
                    sess_objs.append(getattr(the_session.Preferences, attr))
                except Exception:
                    pass
        target_sess_triad = orig_sess_triad if orig_sess_triad is not None else 1
        for s_obj in sess_objs:
            try:
                if hasattr(s_obj, "SetTriadVisibility"):
                    s_obj.SetTriadVisibility(int(target_sess_triad))
            except Exception:
                pass
            try:
                if hasattr(s_obj, "TriadVisibility"):
                    s_obj.TriadVisibility = int(target_sess_triad)
            except Exception:
                pass

    # (D) 工作座標系 (WCS)
    target_wcs = orig_wcs if orig_wcs is not None else True
    if work_part is not None and hasattr(work_part, "WCS"):
        try:
            if hasattr(work_part.WCS, "Visibility"):
                work_part.WCS.Visibility = bool(target_wcs)
        except Exception:
            pass
    if uf_session is not None and hasattr(uf_session, "Csys"):
        try:
            if hasattr(uf_session.Csys, "SetWcsDisplay"):
                uf_session.Csys.SetWcsDisplay(1 if target_wcs else 0)
        except Exception:
            pass

    # (E) 多重強制重繪與刷新圖形緩衝區
    try:
        if w_view is not None and hasattr(w_view, "Update"):
            w_view.Update()
    except Exception:
        pass
    try:
        if uf_session is not None and hasattr(uf_session, "Disp"):
            if hasattr(uf_session.Disp, "RegenerateDisplay"):
                uf_session.Disp.RegenerateDisplay()
            if hasattr(uf_session.Disp, "Refresh"):
                uf_session.Disp.Refresh()
            if hasattr(uf_session.Disp, "MakeDisplayUpToDate"):
                uf_session.Disp.MakeDisplayUpToDate()
    except Exception:
        pass
    try:
        if uf_session is not None and hasattr(uf_session, "View") and w_view is not None:
            if hasattr(uf_session.View, "UpdateView") and hasattr(w_view, "Tag"):
                uf_session.View.UpdateView(w_view.Tag)
    except Exception:
        pass

def get_45deg_craft_orientations():
    """
    定義空間中 26 個以 45° 為單位的標準工藝視向 (以視線法向量定義)
    """
    return [
        # 6 大正交面
        {"name": "Top (正頂俯視)", "vec": (0.0, 0.0, 1.0), "canned": "Top"},
        {"name": "Bottom (正底仰視)", "vec": (0.0, 0.0, -1.0), "canned": "Bottom"},
        {"name": "Front (正前視)", "vec": (0.0, -1.0, 0.0), "canned": "Front"},
        {"name": "Back (正後視)", "vec": (0.0, 1.0, 0.0), "canned": "Back"},
        {"name": "Right (正右視)", "vec": (1.0, 0.0, 0.0), "canned": "Right"},
        {"name": "Left (正左視)", "vec": (-1.0, 0.0, 0.0), "canned": "Left"},

        # 4 個水平 45° 斜平視向
        {"name": "Front-Right 45°", "vec": (0.70710678, -0.70710678, 0.0), "base_canned": "Front", "z_rot": 45.0},
        {"name": "Back-Right 45°", "vec": (0.70710678, 0.70710678, 0.0), "base_canned": "Right", "z_rot": 45.0},
        {"name": "Back-Left 45°", "vec": (-0.70710678, 0.70710678, 0.0), "base_canned": "Back", "z_rot": 45.0},
        {"name": "Front-Left 45°", "vec": (-0.70710678, -0.70710678, 0.0), "base_canned": "Left", "z_rot": 45.0},

        # 4 個 45° 俯視正交棱向
        {"name": "Top-Front 45°", "vec": (0.0, -0.70710678, 0.70710678), "base_canned": "Front", "x_rot": -45.0},
        {"name": "Top-Back 45°", "vec": (0.0, 0.70710678, 0.70710678), "base_canned": "Back", "x_rot": -45.0},
        {"name": "Top-Right 45°", "vec": (0.70710678, 0.0, 0.70710678), "base_canned": "Right", "x_rot": -45.0},
        {"name": "Top-Left 45°", "vec": (-0.70710678, 0.0, 0.70710678), "base_canned": "Left", "x_rot": -45.0},

        # 4 個 45° 俯視等角向 (Isometric 45°)
        {"name": "Isometric Top-Front-Right", "vec": (0.57735027, -0.57735027, 0.57735027), "base_canned": "Front", "combo_rot": (45.0, -35.26)},
        {"name": "Isometric Top-Back-Right", "vec": (0.57735027, 0.57735027, 0.57735027), "base_canned": "Right", "combo_rot": (45.0, -35.26)},
        {"name": "Isometric Top-Back-Left", "vec": (-0.57735027, 0.57735027, 0.57735027), "base_canned": "Back", "combo_rot": (45.0, -35.26)},
        {"name": "Isometric Top-Front-Left", "vec": (-0.57735027, -0.57735027, 0.57735027), "base_canned": "Left", "combo_rot": (45.0, -35.26)},

        # 4 個 45° 仰視正交棱向
        {"name": "Bottom-Front 45°", "vec": (0.0, -0.70710678, -0.70710678), "base_canned": "Front", "x_rot": 45.0},
        {"name": "Bottom-Back 45°", "vec": (0.0, 0.70710678, -0.70710678), "base_canned": "Back", "x_rot": 45.0},
        {"name": "Bottom-Right 45°", "vec": (0.70710678, 0.0, -0.70710678), "base_canned": "Right", "x_rot": 45.0},
        {"name": "Bottom-Left 45°", "vec": (-0.70710678, 0.0, -0.70710678), "base_canned": "Left", "x_rot": 45.0},

        # 4 個 45° 仰視等角向
        {"name": "Isometric Bottom-Front-Right", "vec": (0.57735027, -0.57735027, -0.57735027), "base_canned": "Front", "combo_rot": (45.0, 35.26)},
        {"name": "Isometric Bottom-Back-Right", "vec": (0.57735027, 0.57735027, -0.57735027), "base_canned": "Right", "combo_rot": (45.0, 35.26)},
        {"name": "Isometric Bottom-Back-Left", "vec": (-0.57735027, 0.57735027, -0.57735027), "base_canned": "Back", "combo_rot": (45.0, 35.26)},
        {"name": "Isometric Bottom-Front-Left", "vec": (-0.57735027, -0.57735027, -0.57735027), "base_canned": "Left", "combo_rot": (45.0, 35.26)}
    ]

def snap_work_view_closest(work_part, w_view, uf_session=None):
    """
    底層原生視圖擺正核心 (Snap to Closest 45° Craft Orientation / F8 功能)：
    1. 透過 UFSession 取得當前視圖的 3x3 旋轉矩陣 (AskViewMatrix)
    2. 分析視線法向量 (View Normal)，在 26 個以 45° 為單位的全空間標準工藝視向中尋找最接近之方向
    3. 優先嘗試建立 45° 正交相機矩陣直接貼齊，或先貼齊基準 Canned View 後以 RotateView 精準旋轉 45°
    4. 保證工件世界 Z 軸在螢幕上垂直向上 (無側傾 Roll 偏差)，視角工整整齊！
    5. 即時觸發視圖更新 (Update) 與顯存再生 (RegenerateDisplay)！
    """
    if work_part is None:
        try:
            the_sess = NXOpen.Session.GetSession()
            work_part = the_sess.Parts.Work
        except Exception:
            pass

    if w_view is None and work_part is not None and hasattr(work_part, "Views"):
        try:
            w_view = work_part.Views.WorkView
        except Exception:
            pass

    if w_view is None:
        return False

    orientations = get_45deg_craft_orientations()
    best_cand = orientations[0] # 預設 Top
    fallback_ortho_name = "Top"

    # 1. 嘗試從 UFSession.View.AskViewMatrix 獲取當前視圖矩陣並尋找最接近的 45° 工藝方向
    if uf_session is not None and hasattr(uf_session, "View") and hasattr(w_view, "Tag"):
        try:
            mat = uf_session.View.AskViewMatrix(w_view.Tag)
            if mat and len(mat) >= 9:
                vx, vy, vz = mat[6], mat[7], mat[8]
                v_mag = math.sqrt(vx*vx + vy*vy + vz*vz)
                if v_mag > 1e-6:
                    vx, vy, vz = vx/v_mag, vy/v_mag, vz/v_mag
                else:
                    vx, vy, vz = 0.0, 0.0, 1.0

                # 尋找 26 個 45° 方向中夾角最小 (Dot Product 最大) 者
                best_dot = -999.0
                for cand in orientations:
                    d = vx*cand["vec"][0] + vy*cand["vec"][1] + vz*cand["vec"][2]
                    if d > best_dot:
                        best_dot = d
                        best_cand = cand

                # 同步計算 6 大正交面作為萬一回退之基準
                ortho_dots = {
                    "Top": vz,
                    "Bottom": -vz,
                    "Front": -vy,
                    "Back": vy,
                    "Right": vx,
                    "Left": -vx
                }
                fallback_ortho_name = max(ortho_dots.items(), key=lambda x: x[1])[0]
        except Exception:
            pass

    orient_success = False
    scale_adj = getattr(NXOpen.View.ScaleAdjustment, "Saved", 0)

    # 方案 1 (優先)：若為標準正交面，直接調用原生的 Canned View (如 Top, Front, Right)
    if "canned" in best_cand:
        canned_enum = getattr(NXOpen.View.Canned, best_cand["canned"], None)
        if canned_enum is not None:
            try:
                w_view.Orient(canned_enum, scale_adj)
                orient_success = True
            except Exception:
                pass

    # 方案 2 (次選)：若為 45° 斜角視向且支援 Matrix3x3，建構正交相機矩陣直接貼齊
    if not orient_success:
        try:
            import NXOpen
            N = best_cand["vec"]
            if abs(N[0]) < 1e-4 and abs(N[1]) < 1e-4:
                R = (1.0, 0.0, 0.0)
                U = (0.0, 1.0, 0.0) if N[2] > 0 else (0.0, -1.0, 0.0)
            else:
                Rx = -N[1]
                Ry = N[0]
                Rz = 0.0
                mag_r = math.sqrt(Rx*Rx + Ry*Ry)
                R = (Rx/mag_r, Ry/mag_r, 0.0)
                Ux = N[1]*R[2] - N[2]*R[1]
                Uy = N[2]*R[0] - N[0]*R[2]
                Uz = N[0]*R[1] - N[1]*R[0]
                mag_u = math.sqrt(Ux*Ux + Uy*Uy + Uz*Uz)
                U = (Ux/mag_u, Uy/mag_u, Uz/mag_u)

            new_m = NXOpen.Matrix3x3()
            new_m.Xx, new_m.Xy, new_m.Xz = R[0], R[1], R[2]
            new_m.Yx, new_m.Yy, new_m.Yz = U[0], U[1], U[2]
            new_m.Zx, new_m.Zy, new_m.Zz = N[0], N[1], N[2]

            try:
                w_view.Orient(new_m, scale_adj)
                orient_success = True
            except TypeError:
                try:
                    w_view.Orient(new_m)
                    orient_success = True
                except Exception:
                    w_view.Matrix = new_m
                    orient_success = True
        except Exception:
            pass

    # 方案 3 (三層防護)：若 Orient 矩陣不可用，透過基準 Canned View + RotateView 精確旋轉 45°
    if not orient_success and uf_session is not None and hasattr(uf_session, "View"):
        try:
            base_canned_name = best_cand.get("base_canned") or fallback_ortho_name
            canned_enum = getattr(NXOpen.View.Canned, base_canned_name, None)
            if canned_enum is not None:
                w_view.Orient(canned_enum, scale_adj)
                v_tag = getattr(w_view, "Tag", None)
                if v_tag:
                    if "z_rot" in best_cand:
                        uf_session.View.RotateView(v_tag, 6, best_cand["z_rot"]) # 6 = UF_VIEW_MODEL_Z
                    elif "x_rot" in best_cand:
                        uf_session.View.RotateView(v_tag, 1, best_cand["x_rot"]) # 1 = UF_VIEW_VIEW_X
                    elif "combo_rot" in best_cand:
                        uf_session.View.RotateView(v_tag, 6, best_cand["combo_rot"][0])
                        uf_session.View.RotateView(v_tag, 1, best_cand["combo_rot"][1])
                orient_success = True
        except Exception:
            pass

    # 方案 4 (終極備援)：安全回退至最接近的 6 大正交 Canned View
    if not orient_success:
        try:
            canned_enum = getattr(NXOpen.View.Canned, fallback_ortho_name, None) or getattr(NXOpen.View.Canned, "Top", None)
            if canned_enum:
                w_view.Orient(canned_enum, scale_adj)
                orient_success = True
        except Exception:
            pass

    # 強制視圖重繪刷新
    try:
        if hasattr(w_view, "Update"):
            w_view.Update()
    except Exception:
        pass
    try:
        if uf_session is not None and hasattr(uf_session, "Disp"):
            if hasattr(uf_session.Disp, "RegenerateDisplay"):
                uf_session.Disp.RegenerateDisplay()
            if hasattr(uf_session.Disp, "Refresh"):
                uf_session.Disp.Refresh()
    except Exception:
        pass

    return orient_success

def run_interactive_capture_wizard(stages, temp_dir, uf_session, the_ui=None, listing=None, work_part=None):
    """
    方案二：【無黑窗精緻置頂拍照精靈 + NX 視圖流暢旋轉 + 原生 F8 視角擺正】
    技術核心：
    1. 採用跨機器直譯器探測與純淨環境變數隔離，徹底消除黑窗與直譯器崩潰。
    2. 採用 Popen 非阻塞進程 + Windows 原生訊息泵 (PeekMessage / DispatchMessage)，
       即時泵出滑鼠中鍵與視圖重繪訊息，讓 NX 主視窗在小工具懸浮時 100% 自由流暢旋轉縮放！
    3. 支援【雙重 F8 視角擺正】：
       - 小工具點選【📐 擺正 (F8)】或小工具內按 F8：透過跨進程旗標通知主進程原生執行 snap_work_view_closest
       - NX 主視窗內按 F8：訊息泵即時攔截 WM_KEYDOWN(VK_F8) 直接執行原生擺正！
    """
    if not stages:
        return {}

    captured_images = {}
    total_stages = len(stages)

    # 確保取得當前工作視圖 WorkView
    w_view = None
    if work_part is not None and hasattr(work_part, "Views"):
        try:
            w_view = work_part.Views.WorkView
        except Exception:
            w_view = None

    orig_view_triad = getattr(w_view, "TriadVisibility", None) if w_view else None
    orig_view_wcs = getattr(w_view, "WcsVisibility", None) if w_view else None

    # 1. 尋找精緻置頂拍照組件 capture_assistant_gui.py
    the_session = None
    try:
        the_session = NXOpen.Session.GetSession()
    except Exception:
        pass
    assistant_script = resolve_asset_file("capture_assistant_gui.py", the_session=the_session, work_part=work_part)

    # 2. 取得跨機器純淨 Python 直譯器環境
    interpreter_cmd, clean_env = resolve_python_runtime(require_gui=True)

    gui_success = False

    # 3. 執行置頂拍照小工具 + Windows 訊息泵
    if interpreter_cmd and os.path.exists(assistant_script):
        try:
            import ctypes
            from ctypes import wintypes
            import time

            user32 = ctypes.windll.user32
            msg = wintypes.MSG()
            PM_REMOVE = 0x0001

            for s_idx, stg in enumerate(stages):
                safe_stg_name = re.sub(r'[\\/:*?"<>|]', '_', stg)
                out_img = os.path.join(temp_dir, f"_temp_stage_{safe_stg_name}.png")
                display_name = stg if stg != "通用工段" else "加工工件全貌 / 裝夾示圖"
                req_file = os.path.join(temp_dir, f"_snap_request_{s_idx}.flag")
                if os.path.exists(req_file):
                    try:
                        os.remove(req_file)
                    except Exception:
                        pass

                nx_hwnd = user32.GetForegroundWindow()
                cmd = list(interpreter_cmd) + [
                    assistant_script,
                    "--stage", stg,
                    "--index", str(s_idx + 1),
                    "--total", str(total_stages),
                    "--parent-hwnd", str(nx_hwnd),
                    "--req-file", req_file
                ]

                if listing:
                    listing.WriteLine(f"  [拍照引導] 已啟動右上角精緻引導小按鈕 (請在 NX 中按滑鼠中鍵自由旋轉工件，支援 F8 擺正)...")

                proc = None
                ret_code = -1
                try:
                    # 啟動獨立置頂小工具 (傳入純淨環境 clean_env，無控制台黑窗)
                    proc = subprocess.Popen(
                        cmd,
                        env=clean_env,
                        creationflags=0x08000000
                    )

                    # Windows 原生訊息泵循環：讓 NX 主視窗保持 100% 流暢響應滑鼠中鍵旋轉、重繪與 F8 擺正
                    start_wait = time.time()
                    while proc.poll() is None:
                        # 1. 視角擺正機制 A：檢查來自小工具的擺正旗標請求 (點擊按鈕或小工具焦點按 F8)
                        if os.path.exists(req_file):
                            try:
                                os.remove(req_file)
                            except Exception:
                                pass
                            snap_work_view_closest(work_part, w_view, uf_session)

                        # 2. 視角擺正機制 B：硬體層即時偵測鍵盤 F8 鍵 (VK_F8 = 0x77)
                        # 無論焦點在 NX 主視窗、3D 繪圖區或小工具，只要按 F8 均保證 100% 原生擺正，且不碰訊息隊列
                        try:
                            if user32.GetAsyncKeyState(0x77) & 0x8000:
                                snap_work_view_closest(work_part, w_view, uf_session)
                                time.sleep(0.12) # 消除按鍵重複觸發
                        except Exception:
                            pass

                        # 3. 滑鼠旋轉與畫面重繪：精確派發當前執行緒之滑鼠互動 (0x0200~0x020E) 與重繪訊息 (WM_PAINT 0x000F)
                        # 嚴格限定訊息類型，絕不碰觸 Qt / QtWebEngine 私有 IPC 訊息 (0x8000+ 或 WM_USER)，確保記憶體穩定零洩漏！
                        while user32.PeekMessageW(ctypes.byref(msg), 0, 0x0200, 0x020E, PM_REMOVE):
                            user32.DispatchMessageW(ctypes.byref(msg))
                        while user32.PeekMessageW(ctypes.byref(msg), 0, 0x000F, 0x000F, PM_REMOVE):
                            user32.DispatchMessageW(ctypes.byref(msg))

                        time.sleep(0.015) # 15毫秒平滑睡眠，兼顧滑鼠 60FPS 旋轉流暢感與 CPU 節能

                        # 超時保護 (5分鐘未操作自動退出)
                        if time.time() - start_wait > 300:
                            proc.kill()
                            break

                    ret_code = proc.returncode
                finally:
                    if proc is not None:
                        if proc.poll() is None:
                            try:
                                proc.kill()
                                proc.wait(timeout=2)
                            except Exception:
                                pass
                        try:
                            if proc.stdout: proc.stdout.close()
                            if proc.stderr: proc.stderr.close()
                            if proc.stdin: proc.stdin.close()
                        except Exception:
                            pass

                # 清理旗標檔案
                if os.path.exists(req_file):
                    try:
                        os.remove(req_file)
                    except Exception:
                        pass

                # 狀態代碼精確分流判定：
                if ret_code == 0:
                    # 使用者點選「📸 立即拍照」
                    success = capture_nx_viewport(the_ui, out_img, white_background=True, listing=listing, work_part=work_part)
                    if success:
                        captured_images[stg] = out_img
                        if listing:
                            listing.WriteLine(f"  - 工段【{stg}】示圖已成功截取。")
                elif ret_code == 2:
                    # 使用者在小工具點選「略過此段」或關閉 (X)
                    if listing:
                        listing.WriteLine(f"  - 工段【{stg}】已由使用者略過。")
                else:
                    # 外部小工具異常 (returncode != 0 且 != 2)，自動啟動原生對話框備援，避免功能被跳過
                    if listing:
                        listing.WriteLine(f"  [拍照引導提示] 獨立小工具異常退出 (代碼 {ret_code})，自動啟用 NX 備援確認視窗...")
                    if the_ui:
                        backup_msg = (
                            f"【工段 ({s_idx+1}/{total_stages})】：{display_name}\n\n"
                            f"請確認當前 NX 視窗視角是否已適當？\n\n"
                            f"・點選【是 (Yes)】：立即截取當前視圖 (白底高清)\n"
                            f"・點選【否 (No)】：略過此工段不放圖"
                        )
                        resp = the_ui.NXMessageBox.Show(
                            "CNC 加工示圖拍照確認",
                            NXOpen.NXMessageBox.DialogType.Question,
                            backup_msg
                        )
                        if resp == 1:
                            success = capture_nx_viewport(the_ui, out_img, white_background=True, listing=listing, work_part=work_part)
                            if success:
                                captured_images[stg] = out_img
                                if listing:
                                    listing.WriteLine(f"  - 工段【{stg}】示圖已成功截取。")
                        else:
                            if listing:
                                listing.WriteLine(f"  - 工段【{stg}】已由使用者略過。")

            gui_success = True

        except Exception as ex_pump:
            if listing:
                listing.WriteLine(f"  [拍照引導提示] 獨立小工具啟動異常 ({str(ex_pump)})，自動啟用備援拍照模式...")

    # 4. 備援模式：若外部直譯器異常或未安裝外部 Python，以 NX 原生對話框作為保底
    if not gui_success and the_ui:
        try:
            if listing:
                listing.WriteLine("  [NX 原生引導] 啟動 NX 原生拍照確認視窗引導各工段示圖拍照...")
            for s_idx, stg in enumerate(stages):
                safe_stg_name = re.sub(r'[\\/:*?"<>|]', '_', stg)
                out_img = os.path.join(temp_dir, f"_temp_stage_{safe_stg_name}.png")
                display_name = stg if stg != "通用工段" else "加工工件全貌 / 裝夾示圖"

                dialog_msg = (
                    f"【工段 ({s_idx+1}/{total_stages})】：{display_name}\n\n"
                    f"請在 NX 主視窗旋轉縮放工件至最佳加工示圖角度。\n\n"
                    f"視角確認適當後：\n"
                    f"・點選【是 (Yes)】：立即截取當前視圖 (白底高清)\n"
                    f"・點選【否 (No)】：略過此工段不放圖"
                )
                resp = the_ui.NXMessageBox.Show(
                    "CNC 加工示圖拍照確認",
                    NXOpen.NXMessageBox.DialogType.Question,
                    dialog_msg
                )
                if resp == 1:
                    success = capture_nx_viewport(the_ui, out_img, white_background=True, listing=listing, work_part=work_part)
                    if success:
                        captured_images[stg] = out_img
                        if listing:
                            listing.WriteLine(f"  - 工段【{stg}】示圖已成功截取。")
                else:
                    if listing:
                        listing.WriteLine(f"  - 工段【{stg}】已由使用者略過。")
        except Exception as ex_fb:
            if listing:
                listing.WriteLine(f"  [拍照精靈警告] 備援模式執行異常：{str(ex_fb)}")

    # 拍照流程完全結束，雙重保證 NX 主視窗座標圖示 (View Triad / WCS) 100% 恢復顯示並強制刷新
    try:
        restore_triad_and_wcs(
            the_session=None, work_part=work_part, w_view=w_view, uf_session=uf_session,
            orig_view_triad=orig_view_triad, orig_view_wcs=orig_view_wcs
        )
    except Exception:
        pass

    return captured_images

# ==================== 分頁排版核心模組 (Pagination Engine) ====================

def paginate_operations(chunks, rows_per_page=ROWS_PER_PAGE, pagination_mode="paginate"):
    """
    核心分頁演算法 (支援工段優先分頁與超過 15 格單頁延伸/截圖下一頁模式)：
    1. 工段優先決定分頁 (跨工段絕對強制分頁，各工段版面獨立)
    2. 標準分頁模式 (pagination_mode == "paginate")：
       - 每頁上限為 rows_per_page (預設 10 列)
       - 每頁皆於 A18:J36 嵌入加工示圖
    3. 單頁延伸模式 (pagination_mode == "extend")：
       - 若該工段工步數 <= 15：維持標準分頁
       - 若 15 < 總工步數 <= 20：第一頁刀具向下延伸，下方示圖縮小置於 Row (8+N) ~ 36
       - 若 總工步數 > 20：刀具延伸填滿第一頁 (最多30格)，截圖放不下則自動顯示於下一頁 (專屬大示圖頁 A7:J36)
    回傳：[{"stage": stage_key, "rows": page_rows, "show_image": bool, "image_top_row": int, "image_bottom_row": int, "is_appendix_image_page": bool}, ...]
    """
    pages = []

    # 1. 依工段 (Stage) 先行聚合 Chunks，確保工段優先決定分頁
    stage_groups = {}
    stage_order = []
    for chunk in chunks:
        grp_name = chunk["group_name"]
        stg = chunk.get("stage") or extract_stage_key(grp_name) or "通用工段"
        if stg not in stage_groups:
            stage_groups[stg] = []
            stage_order.append(stg)
        stage_groups[stg].append(chunk)

    for stg in stage_order:
        chunks_in_stage = stage_groups[stg]
        all_ops = []
        for c in chunks_in_stage:
            for op_item in c.get("operations", []):
                all_ops.append(op_item)

        total_ops = len(all_ops)
        if total_ops <= 15:
            # === 工步數在 1 ~ 15 格以內：單頁自動增加格數容納，絕不切頁！ ===
            page_rows = [{"type": "op", "data": op} for op in all_ops]
            img_top = 18 if total_ops <= 10 else (8 + total_ops)
            pages.append({
                "stage": stg,
                "rows": page_rows,
                "show_image": True,
                "image_top_row": img_top,
                "image_bottom_row": 36,
                "is_appendix_image_page": False
            })
        else:
            # === 工步數超過 15 格：依使用者選擇決定 ===
            if pagination_mode == "extend":
                # 單頁延伸模式
                if total_ops <= 20:
                    # 狀況 A：16 ~ 20 步，下方空間足夠，截圖縮小放於同頁下方
                    page_rows = [{"type": "op", "data": op} for op in all_ops]
                    img_top = 8 + total_ops
                    pages.append({
                        "stage": stg,
                        "rows": page_rows,
                        "show_image": True,
                        "image_top_row": img_top,
                        "image_bottom_row": 36,
                        "is_appendix_image_page": False
                    })
                else:
                    # 狀況 B：超過 20 步，影響到截圖顯示，截圖移至下一頁！
                    # 第 1 頁：刀具延伸頁 (最多 30 格)
                    first_page_ops = all_ops[:30]
                    pages.append({
                        "stage": stg,
                        "rows": [{"type": "op", "data": op} for op in first_page_ops],
                        "show_image": False,
                        "image_top_row": None,
                        "image_bottom_row": None,
                        "is_appendix_image_page": False
                    })

                    # 若超過 30 步 (極少見)，續頁放刀具
                    rem_ops = all_ops[30:]
                    while rem_ops:
                        chunk_ops = rem_ops[:30]
                        rem_ops = rem_ops[30:]
                        pages.append({
                            "stage": stg,
                            "rows": [{"type": "op", "data": op} for op in chunk_ops],
                            "show_image": False,
                            "image_top_row": None,
                            "image_bottom_row": None,
                            "is_appendix_image_page": False
                        })

                    # 專屬示圖頁：表頭相同，中間為清晰大示圖 (A7:J36)
                    pages.append({
                        "stage": stg,
                        "rows": [],  # 刀具列留空
                        "show_image": True,
                        "image_top_row": 7,
                        "image_bottom_row": 36,
                        "is_appendix_image_page": True
                    })
            else:
                # === 標準分頁模式 (使用者選擇分頁：每頁 10 格，每頁含示圖) ===
                current_page = []
                for op_item in all_ops:
                    if len(current_page) >= rows_per_page:
                        pages.append({
                            "stage": stg,
                            "rows": current_page,
                            "show_image": True,
                            "image_top_row": 18,
                            "image_bottom_row": 36,
                            "is_appendix_image_page": False
                        })
                        current_page = []
                    current_page.append({"type": "op", "data": op_item})

                if current_page:
                    pages.append({
                        "stage": stg,
                        "rows": current_page,
                        "show_image": True,
                        "image_top_row": 18,
                        "image_bottom_row": 36,
                        "is_appendix_image_page": False
                    })

    return pages

def preview_pagination_plan(processed_chunks, rows_per_page=ROWS_PER_PAGE):
    """
    預先分析各工段工步數量，並精確計算「單頁延伸」與「自動分頁」模式下的各工段頁數與總頁數。
    回傳：(stage_stats, total_extend_pages, total_paginate_pages, has_choice)
    """
    stage_groups = {}
    stage_order = []
    for chunk in processed_chunks:
        grp_name = chunk["group_name"]
        stg = chunk.get("stage") or extract_stage_key(grp_name) or "通用工段"
        if stg not in stage_groups:
            stage_groups[stg] = []
            stage_order.append(stg)
        stage_groups[stg].append(chunk)

    stage_stats = []
    total_extend_pages = 0
    total_paginate_pages = 0
    has_choice = False

    for stg in stage_order:
        chunks_in_stage = stage_groups[stg]
        all_ops = []
        for c in chunks_in_stage:
            all_ops.extend(c.get("operations", []))
        cnt = len(all_ops)

        # 1. 計算單頁延伸模式頁數
        if cnt <= 20:
            ext_p = 1
            if cnt <= 10:
                ext_desc = "第1頁 (標準10格版面，含完整示圖)"
            elif cnt <= 15:
                ext_desc = f"第1頁 (自動增加至 {cnt} 格，同頁微縮示圖)"
            else:
                ext_desc = f"第1頁 (向下延伸至 {cnt} 格，同頁微縮示圖)"
        else:
            # > 20 刀：放不下示圖，刀具延伸整頁 (最多30刀/頁) + 專屬大示圖 1 頁
            tool_pages = max(1, (cnt + 29) // 30)
            ext_p = tool_pages + 1
            ext_desc = f"共 {ext_p} 頁 (第1頁延伸 {min(30, cnt)} 刀，第 {ext_p} 頁專屬大示圖)"

        # 2. 計算自動分頁模式頁數
        pag_p = max(1, (cnt + rows_per_page - 1) // rows_per_page)
        if pag_p == 1:
            pag_desc = "第1頁 (標準10格版面，含完整示圖)"
        else:
            rem = cnt % rows_per_page or rows_per_page
            pag_desc = f"共 {pag_p} 頁 (每頁 {rows_per_page} 刀，末頁 {rem} 刀，每頁皆含示圖)"

        if cnt > 10:
            has_choice = True

        stage_stats.append({
            "stage": stg,
            "count": cnt,
            "extend_pages": ext_p,
            "paginate_pages": pag_p,
            "extend_desc": ext_desc,
            "paginate_desc": pag_desc
        })

        total_extend_pages += ext_p
        total_paginate_pages += pag_p

    return stage_stats, total_extend_pages, total_paginate_pages, has_choice

# ==================== VBS 多頁動態生成與匯出模組 ====================

def escape_vbs_str(val):
    """
    將字串安全轉義為 VBScript 字串常數：
    1. 雙引號轉義為兩個雙引號 (" -> "")
    2. 清除換行符號 (\r, \n)，轉為空格，防止 VBScript 跨行語法中斷 (Unterminated string constant)
    """
    if val is None:
        return ""
    s = str(val).replace('"', '""')
    s = s.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
    return s

def export_multipage_via_vbs(pages, template_path, output_path, work_part, header_info, stage_images=None):
    """
    透過 Windows 原生 VBScript 動態複製 Excel 工作表產生「第1頁」、「第2頁」...
    並填入表頭、工步明細、頁碼標註，並將工段對應之加工示圖等比例居中嵌入 A18:J36 區域
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到工單範本：{template_path}")

    stage_images = stage_images or {}
    part_name = work_part.Leaf
    part_path = work_part.FullPath
    part_dir = os.path.dirname(part_path)
    today_str = datetime.datetime.now().strftime("%Y/%m/%d")

    # 提取表頭資訊 (依使用者最新規則，經過 escape_vbs_str 嚴格過濾雙引號與換行符)
    # 1. 圖號：若檔案名稱當中有 14 碼編號則為圖號，若無則留空
    drawing_number = escape_vbs_str(header_info.get("drawing_number", ""))
    # 2. 圖名：應為檔案名稱
    drawing_name = escape_vbs_str(header_info.get("drawing_name", part_name))
    # 3. 尺寸：由設定的素材大小決定，若無資訊則留空
    blank_size = escape_vbs_str(header_info.get("blank_size", ""))

    part_no_val = escape_vbs_str(header_info.get("part_number", "")) # 工單編號/料號
    holes_val = escape_vbs_str(header_info.get("holes", ""))         # 數量/孔數

    part_path_vbs = escape_vbs_str(part_path)
    part_dir_vbs = escape_vbs_str(part_dir)
    abs_template = escape_vbs_str(os.path.abspath(template_path))
    abs_output = escape_vbs_str(os.path.abspath(output_path))
    total_pages = len(pages) if len(pages) > 0 else 1

    vbs_lines = [
        'Dim objExcel, objWb, seedWs, ws, fso',
        'Dim shp, topCell, bottomCell, boxL, boxT, boxW, boxH, origW, origH, targetW, targetH',
        'Dim r, vbsErrCode, vbsErrDesc',
        'vbsErrCode = 0',
        'vbsErrDesc = ""',
        'On Error Resume Next',
        'Set objExcel = CreateObject("Excel.Application")',
        'If Err.Number <> 0 Then',
        '    WScript.StdErr.WriteLine "無法啟動 Excel.Application: " & Err.Description',
        '    WScript.Quit Err.Number',
        'End If',
        'objExcel.Visible = False',
        'objExcel.DisplayAlerts = False',
        'Set fso = CreateObject("Scripting.FileSystemObject")',
        f'Set objWb = objExcel.Workbooks.Open("{abs_template}")',
        'If Err.Number <> 0 Then',
        '    vbsErrCode = Err.Number',
        '    vbsErrDesc = "無法開啟工單範本: " & Err.Description',
        '    objExcel.Quit',
        '    Set objExcel = Nothing',
        '    WScript.StdErr.WriteLine vbsErrDesc',
        '    WScript.Quit vbsErrCode',
        'End If',
        'Set seedWs = objWb.Sheets(1)'
    ]

    # 逐頁複製母版並填入資料
    for p_idx, page_item in enumerate(pages, start=1):
        page_name = f"第{p_idx}頁"
        page_str = f"{p_idx}-{total_pages}"

        if isinstance(page_item, dict):
            page_rows = page_item.get("rows", [])
            page_stage = page_item.get("stage", "通用工段")
            show_image = page_item.get("show_image", True)
            image_top_row = page_item.get("image_top_row", 18)
            image_bottom_row = page_item.get("image_bottom_row", 36)
            is_appendix = page_item.get("is_appendix_image_page", False)
        else:
            page_rows = page_item
            page_stage = "通用工段"
            show_image = True
            image_top_row = 18
            image_bottom_row = 36
            is_appendix = False

        # 複製母版工作表至尾端
        vbs_lines.extend([
            'seedWs.Copy , objWb.Sheets(objWb.Sheets.Count)',
            'Set ws = objWb.Sheets(objWb.Sheets.Count)',
            f'ws.Name = "{page_name}"',
            # 填入表頭基本資訊 (精確對應 ShopDoc_Template.xlsx 欄位定義)
            f'ws.Range("B1").Value = "{drawing_number}"',  # A1「圖號」 (14碼/留空)
            f'ws.Range("B2").Value = "{drawing_name}"',    # A2「圖名」 (檔案名稱)
            f'ws.Range("G1").Value = "{part_no_val}"',     # E1「工單編號」
            f'ws.Range("B3").Value = "{blank_size}"',      # A3「尺寸」 (素材大小/留空)
            f'ws.Range("D3").Value = "{holes_val}"',       # C3「數量」
            f'ws.Range("J2").Value = "{today_str}"',       # I2「表單日期」
            f'ws.Range("B4").Value = "{part_path_vbs}"',       # A4「檔案位置」
            f'ws.Range("B5").Value = "{part_dir_vbs}"',        # A5「程式位置」
            'ws.Range("J37").NumberFormat = "@"',          # 強制指定頁數欄位為純文字格式，防止 Excel 自動轉換為日期 (如 1-7 變成 1月7日)
            f'ws.Range("J37").Value = "{page_str}"'        # 頁數標記 (J37)
        ])

        start_row = 7
        page_seq = 1  # 每頁有效工步流水號，從 1 開始計算 (空行不計)
        num_rows = len(page_rows)

        if is_appendix:
            # 專屬示圖頁：清空 Row 7 ~ Row 16 預設欄位與邊框，讓上方乾淨開闊
            vbs_lines.extend([
                'ws.Range("A7:J16").ClearContents',
                'ws.Range("A7:J16").Borders.LineStyle = -4142'
            ])
        else:
            # 若刀具列超過 10 列 (自動增加格數 / 單頁延伸模式)，向下精確複製 A16:J16 格式與合併格
            if num_rows > ROWS_PER_PAGE:
                last_op_row = start_row + num_rows - 1
                vbs_lines.extend([
                    f'For r = 17 To {last_op_row}',
                    '    ws.Range("A16:J16").Copy',
                    '    ws.Range("A" & r & ":J" & r).PasteSpecial -4122',
                    'Next',
                    'objExcel.CutCopyMode = False',
                    f'ws.Range("A{last_op_row}:J{last_op_row}").Borders(4).Weight = 3'  # 4 = xlEdgeBottom, 3 = xlMedium 封底線
                ])

            loop_rows = max(ROWS_PER_PAGE, num_rows)
            for row_offset in range(loop_rows):
                curr_r = start_row + row_offset
                if row_offset < num_rows:
                    item = page_rows[row_offset]
                    if item["type"] == "spacer":
                        # 空行分隔：清空此列資料 (不計序號)
                        for col_i in range(1, 9):
                            vbs_lines.append(f'ws.Cells({curr_r}, {col_i}).Value = ""')
                    else:
                        d = item["data"]
                        safe_seq = str(page_seq)
                        safe_op_name = escape_vbs_str(d.get("op_name", ""))
                        safe_tool_num = escape_vbs_str(d.get("tool_number", ""))
                        spec_str = d.get("tool_spec_display", d.get("tool_diameter", "-"))
                        safe_tool_dia = escape_vbs_str(spec_str)
                        safe_flute_len = escape_vbs_str(d.get("flute_length", ""))
                        safe_holder_len = escape_vbs_str(d.get("holder_length", ""))
                        safe_time = escape_vbs_str(d.get("time", ""))
                        safe_note = escape_vbs_str(d.get("note", ""))

                        vbs_lines.append(f'ws.Cells({curr_r}, 1).Value = "{safe_seq}"')
                        vbs_lines.append(f'ws.Cells({curr_r}, 2).Value = "{safe_op_name}"')
                        vbs_lines.append(f'ws.Cells({curr_r}, 3).Value = "{safe_tool_num}"')
                        vbs_lines.append(f'ws.Cells({curr_r}, 4).Value = "{safe_tool_dia}"')
                        vbs_lines.append(f'ws.Cells({curr_r}, 5).Value = "{safe_flute_len}"')
                        vbs_lines.append(f'ws.Cells({curr_r}, 6).Value = "{safe_holder_len}"')
                        vbs_lines.append(f'ws.Cells({curr_r}, 7).Value = "{safe_time}"')
                        vbs_lines.append(f'ws.Cells({curr_r}, 8).Value = "{safe_note}"')
                        page_seq += 1
                else:
                    # 未填滿的列位：清空範本預設的工站數字，保持頁面乾淨
                    for col_i in range(1, 9):
                        vbs_lines.append(f'ws.Cells({curr_r}, {col_i}).Value = ""')

        # 插入該工段對應之加工示圖 (等比例居中置放)
        img_path = stage_images.get(page_stage)
        if not img_path:
            img_path = stage_images.get("通用工段")
        if not img_path and len(stage_images) == 1:
            img_path = list(stage_images.values())[0]

        if show_image and img_path and os.path.exists(img_path):
            safe_img = os.path.abspath(img_path).replace('"', '""')
            top_cell_str = f"A{image_top_row}"
            bottom_cell_str = f"J{image_bottom_row}"
            vbs_lines.extend([
                f'If fso.FileExists("{safe_img}") Then',
                f'    Set topCell = ws.Range("{top_cell_str}")',
                f'    Set bottomCell = ws.Range("{bottom_cell_str}")',
                '    boxL = topCell.Left + 5',
                '    boxT = topCell.Top + 5',
                '    boxW = (bottomCell.Left + bottomCell.Width) - topCell.Left - 10',
                '    boxH = (bottomCell.Top + bottomCell.Height) - topCell.Top - 10',
                f'    Set shp = ws.Shapes.AddPicture("{safe_img}", 0, -1, boxL, boxT, -1, -1)',
                '    shp.LockAspectRatio = -1',
                '    origW = shp.Width',
                '    origH = shp.Height',
                '    If (origW / origH) > (boxW / boxH) Then',
                '        targetW = boxW',
                '        targetH = origH * (boxW / origW)',
                '    Else',
                '        targetH = boxH',
                '        targetW = origW * (boxH / origH)',
                '    End If',
                '    shp.Width = targetW',
                '    shp.Height = targetH',
                '    shp.Left = boxL + (boxW - targetW) / 2',
                '    shp.Top = boxT + (boxH - targetH) / 2',
                '    shp.Placement = 1',
                'End If'
            ])

    # 刪除最初的母版，只保留產出的 第1頁, 第2頁...
    vbs_lines.extend([
        'seedWs.Delete',
        f'objWb.SaveAs "{abs_output}"',
        'If Err.Number <> 0 Then',
        '    vbsErrCode = Err.Number',
        '    vbsErrDesc = "Excel 寫入或儲存錯誤: " & Err.Description',
        'End If',
        'If Not objWb Is Nothing Then',
        '    objWb.Close False',
        'End If',
        'If Not objExcel Is Nothing Then',
        '    objExcel.Quit',
        'End If',
        'Set shp = Nothing',
        'Set topCell = Nothing',
        'Set bottomCell = Nothing',
        'Set ws = Nothing',
        'Set seedWs = Nothing',
        'Set objWb = Nothing',
        'Set objExcel = Nothing',
        'Set fso = Nothing',
        'If vbsErrCode <> 0 Then',
        '    WScript.StdErr.WriteLine vbsErrDesc',
        '    WScript.Quit vbsErrCode',
        'End If'
    ])

    vbs_content = "\r\n".join(vbs_lines)
    temp_dir = tempfile.gettempdir()
    temp_vbs = os.path.join(temp_dir, f"_temp_shopdoc_{os.getpid()}_{datetime.datetime.now().strftime('%H%M%S%f')}.vbs")

    with open(temp_vbs, "w", encoding="cp950", errors="ignore") as f:
        f.write(vbs_content)

    try:
        subprocess.run(
            ["cscript.exe", "//Nologo", temp_vbs],
            check=True,
            timeout=180,
            capture_output=True,
            text=True,
            creationflags=0x08000000
        )
    except subprocess.TimeoutExpired:
        raise TimeoutError("Excel 匯出程序執行超時 (超過 180 秒)，已強制終止進程，避免記憶體與程序卡死。")
    except subprocess.CalledProcessError as cpe:
        err_detail = cpe.stderr.strip() if cpe.stderr else cpe.stdout.strip()
        raise RuntimeError(f"Excel 匯出失敗 (VBScript 代碼 {cpe.returncode})：{err_detail}")
    finally:
        if os.path.exists(temp_vbs):
            try:
                os.remove(temp_vbs)
            except Exception:
                pass

# ==================== NX CAM 後處理與 NC 轉出模組 ====================

def build_nc_tasks_from_selection(selected_objects):
    """
    依照使用者在 CAM 導覽器選取之物件構建 NC 轉出任務列表：
    1. 若選取為父資料夾：將該資料夾內所有底層工序彙整為一個 NC 任務，檔名為資料夾名稱 (如 MK-1-M1.nc)。
    2. 若祖先已被選取：該子資料夾或子工序自動略過（由最頂層選取統一轉出，避免重疊重複產出）。
    3. 若單獨選取工序：該工序獨立為一個 NC 任務，檔名為該工序名稱。
    4. 保證 100% 遵守選取之工藝順序與去重原則。
    """
    tasks = []
    covered_op_tags = set()
    selected_set = set(selected_objects)

    def _has_selected_ancestor(node):
        curr = node
        while curr is not None:
            try:
                curr = curr.GetParent(NXOpen.CAM.CAMSetup.View.ProgramOrder)
            except Exception:
                curr = None
            if curr and curr in selected_set:
                return True
        return False

    def _collect_ops_in_group(grp):
        ops = []
        name = getattr(grp, "Name", "").strip()
        if is_excluded_group_name(name) or check_and_extract_info(name, {}):
            return ops
        try:
            for member in grp.GetMembers():
                if isinstance(member, NXOpen.CAM.Operation):
                    ops.append(member)
                elif isinstance(member, NXOpen.CAM.NCGroup):
                    ops.extend(_collect_ops_in_group(member))
        except Exception:
            pass
        return ops

    for obj in selected_objects:
        obj_name = getattr(obj, "Name", "").strip()
        is_grp = isinstance(obj, NXOpen.CAM.NCGroup)

        # 檢查是否已有祖先被選取，若有則此節點已被納入祖先任務，略過
        if _has_selected_ancestor(obj):
            continue

        if is_grp:
            if is_excluded_group_name(obj_name) or check_and_extract_info(obj_name, {}):
                continue
            group_ops = _collect_ops_in_group(obj)
            # 濾除已涵蓋之工序
            remaining_ops = [op for op in group_ops if getattr(op, "Tag", None) not in covered_op_tags]
            if remaining_ops:
                clean_name = clean_program_name(obj_name)
                clean_name = re.sub(r'[\\/:*?"<>|]', '_', clean_name)
                tasks.append({
                    "program_name": clean_name,
                    "operations": remaining_ops,
                    "op_names": [getattr(op, "Name", "") for op in remaining_ops],
                    "op_count": len(remaining_ops),
                    "source_type": "group"
                })
                for op in remaining_ops:
                    t = getattr(op, "Tag", None)
                    if t:
                        covered_op_tags.add(t)
        elif isinstance(obj, NXOpen.CAM.Operation):
            # 單獨工序
            op_tag = getattr(obj, "Tag", None)
            if op_tag not in covered_op_tags:
                clean_name = clean_program_name(obj_name)
                clean_name = re.sub(r'[\\/:*?"<>|]', '_', clean_name)
                tasks.append({
                    "program_name": clean_name,
                    "operations": [obj],
                    "op_names": [obj_name],
                    "op_count": 1,
                    "source_type": "operation"
                })
                if op_tag:
                    covered_op_tags.add(op_tag)

    return tasks

def invoke_nc_post_dialog(nc_tasks, default_dir, default_post="Fanuc_2026", listing=None, the_session=None, work_part=None):
    """
    啟動獨立進程之後處理確認視窗 (nc_post_dialog.py)
    回傳 post_config 字典：
      {
         "action": "post_and_export" | "export_only" | "cancel",
         "postprocessor_name": "Fanuc_2026",
         "custom_post_path": "",
         "output_dir": r"...",
         "extension": ".nc"
      }
    """
    if the_session is None:
        try:
            the_session = NXOpen.Session.GetSession()
        except Exception:
            pass

    dialog_script = resolve_asset_file("nc_post_dialog.py", the_session=the_session, work_part=work_part)
    interpreter_cmd, clean_env = resolve_python_runtime(require_gui=True)

    if not interpreter_cmd or not os.path.exists(dialog_script):
        # 備援防護：若無外部 Python 直譯器或獨立組件，啟動 NX 原生對話框完成決策，確保純 NX 電腦依然具備 100% 決策能力
        try:
            the_ui = NXOpen.UI.GetUI()
            if the_ui:
                task_names = [t.get("program_name", "") for t in nc_tasks if t.get("program_name", "")]
                task_summary_str = "、".join(task_names[:5])
                if len(task_names) > 5:
                    task_summary_str += f" 等共 {len(task_names)} 組"

                msg = (
                    "【後處理轉出確認 (NX 原生備援視窗)】\n\n"
                    f"選定之 NC 程式：{task_summary_str}\n"
                    f"預設後處理器：{default_post}\n"
                    f"輸出目錄：{default_dir}\n\n"
                    "是否要依照選定工序自動轉出 NC 碼？\n\n"
                    "・點選【是 (Yes)】：轉出 NC 碼並匯出工單\n"
                    "・點選【否 (No)】：僅匯出工單 (不轉出 NC 碼)"
                )
                resp = the_ui.NXMessageBox.Show(
                    "後處理轉出與工單設定",
                    NXOpen.NXMessageBox.DialogType.Question,
                    msg
                )
                if resp == 1:
                    if listing:
                        listing.WriteLine("  [NX 備援視窗] 使用者選擇【轉出 NC 碼並匯出工單】。")
                    return {
                        "action": "post_and_export",
                        "postprocessor_name": default_post,
                        "custom_post_path": "",
                        "output_dir": default_dir,
                        "extension": ".nc"
                    }
                else:
                    if listing:
                        listing.WriteLine("  [NX 備援視窗] 使用者選擇【僅匯出工單 (不轉 NC 碼)】。")
                    return {"action": "export_only"}
        except Exception:
            pass

        if listing:
            if not interpreter_cmd:
                listing.WriteLine("  [提示] 未能偵測到支援 Tkinter 之 Python 直譯器，預設僅匯出工單。")
            elif not os.path.exists(dialog_script):
                listing.WriteLine(f"  [提示] 找不到後處理對話框組件 ({dialog_script})，預設僅匯出工單。")
            else:
                listing.WriteLine("  [提示] 找不到外部 Python 或對話框組件，預設僅匯出工單。")
        return {"action": "export_only"}

    temp_dir = tempfile.gettempdir()
    pid = os.getpid()
    timestamp = datetime.datetime.now().strftime("%H%M%S%f")
    cfg_file = os.path.join(temp_dir, f"_nc_post_cfg_{pid}_{timestamp}.json")
    res_file = os.path.join(temp_dir, f"_nc_post_res_{pid}_{timestamp}.json")

    # 序列化任務資料 (過濾 NX 原生物件)
    tasks_summary = []
    for t in nc_tasks:
        tasks_summary.append({
            "program_name": t.get("program_name", ""),
            "op_count": t.get("op_count", 0),
            "op_names": t.get("op_names", []),
            "source_type": t.get("source_type", "")
        })

    config_payload = {
        "output_dir": default_dir,
        "default_post": default_post,
        "nc_tasks": tasks_summary
    }

    try:
        with open(cfg_file, "w", encoding="utf-8") as f:
            json.dump(config_payload, f, ensure_ascii=False, indent=2)
    except Exception as ex:
        if listing:
            listing.WriteLine(f"  [錯誤] 無法寫入後處理設定暫存檔：{str(ex)}")
        return {"action": "export_only"}

    cmd = list(interpreter_cmd) + [
        dialog_script,
        "--cfg-file", cfg_file,
        "--res-file", res_file
    ]

    result = {"action": "export_only"}
    proc = None
    try:
        proc = subprocess.Popen(
            cmd,
            env=clean_env,
            creationflags=0x08000000
        )

        # 安全等待外部對話框完成 (絕對不呼叫 PeekMessage 搶奪 Qt / QtWebEngine 私有訊息隊列)
        try:
            proc.wait(timeout=600) # 10 分鐘保護
        except subprocess.TimeoutExpired:
            proc.kill()

        if os.path.exists(res_file):
            with open(res_file, "r", encoding="utf-8") as f:
                result = json.load(f)
    except Exception as ex:
        if listing:
            listing.WriteLine(f"  [後處理對話框例外] {str(ex)}，自動啟用 NX 原生備援視窗...")
        try:
            the_ui = NXOpen.UI.GetUI()
            if the_ui:
                task_names = [t.get("program_name", "") for t in nc_tasks if t.get("program_name", "")]
                task_summary_str = "、".join(task_names[:5])
                if len(task_names) > 5:
                    task_summary_str += f" 等共 {len(task_names)} 組"

                msg = (
                    "【後處理轉出確認 (NX 原生備援視窗)】\n\n"
                    f"選定之 NC 程式：{task_summary_str}\n"
                    f"預設後處理器：{default_post}\n"
                    f"輸出目錄：{default_dir}\n\n"
                    "是否要依照選定工序自動轉出 NC 碼？\n\n"
                    "・點選【是 (Yes)】：轉出 NC 碼並匯出工單\n"
                    "・點選【否 (No)】：僅匯出工單 (不轉出 NC 碼)"
                )
                resp = the_ui.NXMessageBox.Show(
                    "後處理轉出與工單設定",
                    NXOpen.NXMessageBox.DialogType.Question,
                    msg
                )
                if resp == 1:
                    result = {
                        "action": "post_and_export",
                        "postprocessor_name": default_post,
                        "custom_post_path": "",
                        "output_dir": default_dir,
                        "extension": ".nc"
                    }
                else:
                    result = {"action": "export_only"}
        except Exception:
            result = {"action": "export_only"}
    finally:
        if proc is not None:
            if proc.poll() is None:
                try:
                    proc.kill()
                    proc.wait(timeout=2)
                except Exception:
                    pass
            try:
                if proc.stdout: proc.stdout.close()
                if proc.stderr: proc.stderr.close()
                if proc.stdin: proc.stdin.close()
            except Exception:
                pass
        for tmp_f in [cfg_file, res_file]:
            if os.path.exists(tmp_f):
                try:
                    os.remove(tmp_f)
                except Exception:
                    pass

    return result

def invoke_pagination_prompt_dialog(stage_stats, total_extend_pages, total_paginate_pages, listing=None, the_session=None, work_part=None):
    """
    啟動獨立進程之工單分頁預覽與排版決策對話視窗 (pagination_prompt_dialog.py)
    回傳: "extend" (單頁延伸顯示), "paginate" (自動分頁顯示) 或 "cancel" (取消操作)
    """
    if the_session is None:
        try:
            the_session = NXOpen.Session.GetSession()
        except Exception:
            pass

    dialog_script = resolve_asset_file("pagination_prompt_dialog.py", the_session=the_session, work_part=work_part)
    interpreter_cmd, clean_env = resolve_python_runtime(require_gui=True)

    if not interpreter_cmd or not os.path.exists(dialog_script):
        # 備援防護：若無外部 Python 直譯器或獨立組件，啟動 NX 原生對話框完成決策，確保純 NX 電腦依然具備 100% 決策能力
        try:
            the_ui = NXOpen.UI.GetUI()
            if the_ui:
                stage_info_lines = []
                for s in stage_stats:
                    stage_info_lines.append(f"  ・工段【{s['stage']}】({s['count']}刀)：延伸 ➔ {s['extend_desc']} | 分頁 ➔ {s['paginate_desc']}")
                stage_desc_str = "\n".join(stage_info_lines[:4])

                msg = (
                    "【工單分頁預覽與排版決策 (NX 原生備援視窗)】\n\n"
                    f"目前工段刀具數已超過 15 格：\n"
                    f"{stage_desc_str}\n\n"
                    f"預計總頁數對比：\n"
                    f"  - 單頁延伸模式：共 {total_extend_pages} 頁 (刀具列向下增加格數)\n"
                    f"  - 自動分頁模式：共 {total_paginate_pages} 頁 (標準 10 格/頁分頁)\n\n"
                    "請選擇排版模式：\n"
                    "・點選【是 (Yes)】：自動分頁顯示 (共 " + str(total_paginate_pages) + " 頁)\n"
                    "・點選【否 (No)】：單頁延伸顯示 (共 " + str(total_extend_pages) + " 頁)"
                )
                resp = the_ui.NXMessageBox.Show(
                    "工單分頁模式決策",
                    NXOpen.NXMessageBox.DialogType.Question,
                    msg
                )
                if resp == 1:
                    if listing:
                        listing.WriteLine(f"  [NX 備援視窗] 使用者選擇【自動分頁顯示】(共 {total_paginate_pages} 頁)。")
                    return "paginate"
                else:
                    if listing:
                        listing.WriteLine(f"  [NX 備援視窗] 使用者選擇【單頁延伸顯示】(共 {total_extend_pages} 頁)。")
                    return "extend"
        except Exception:
            pass

        if listing:
            if not interpreter_cmd:
                listing.WriteLine("  [提示] 未能偵測到支援 Tkinter 之 Python 直譯器，預設採用單頁延伸排版模式。")
            elif not os.path.exists(dialog_script):
                listing.WriteLine(f"  [提示] 找不到分頁對話框組件 ({dialog_script})，預設採用單頁延伸排版模式。")
            else:
                listing.WriteLine("  [提示] 找不到外部 Python 或對話框組件，預設採用單頁延伸排版模式。")
        return "extend"

    temp_dir = tempfile.gettempdir()
    pid = os.getpid()
    timestamp = datetime.datetime.now().strftime("%H%M%S%f")
    cfg_file = os.path.join(temp_dir, f"_p_cfg_{pid}_{timestamp}.json")
    res_file = os.path.join(temp_dir, f"_p_res_{pid}_{timestamp}.json")

    config_payload = {
        "stage_stats": stage_stats,
        "total_extend_pages": total_extend_pages,
        "total_paginate_pages": total_paginate_pages
    }

    try:
        with open(cfg_file, "w", encoding="utf-8") as f:
            json.dump(config_payload, f, ensure_ascii=False, indent=2)
    except Exception as ex:
        if listing:
            listing.WriteLine(f"  [錯誤] 無法寫入分頁設定暫存檔：{str(ex)}")
        return "extend"

    cmd = list(interpreter_cmd) + [
        dialog_script,
        "--cfg-file", cfg_file,
        "--res-file", res_file
    ]

    mode_result = "extend"
    proc = None
    try:
        proc = subprocess.Popen(
            cmd,
            env=clean_env,
            creationflags=0x08000000
        )

        # 安全等待外部對話框完成 (絕對不呼叫 PeekMessage 搶奪 Qt / QtWebEngine 私有訊息隊列)
        try:
            proc.wait(timeout=300) # 5 分鐘保護
        except subprocess.TimeoutExpired:
            proc.kill()

        if os.path.exists(res_file):
            with open(res_file, "r", encoding="utf-8") as f:
                res_data = json.load(f)
                mode_result = res_data.get("mode", "extend")
    except Exception as ex:
        if listing:
            listing.WriteLine(f"  [分頁對話框例外] {str(ex)}，自動啟用 NX 原生備援視窗...")
        try:
            the_ui = NXOpen.UI.GetUI()
            if the_ui:
                stage_info_lines = []
                for s in stage_stats:
                    stage_info_lines.append(f"  ・工段【{s['stage']}】({s['count']}刀)：延伸 ➔ {s['extend_desc']} | 分頁 ➔ {s['paginate_desc']}")
                stage_desc_str = "\n".join(stage_info_lines[:4])

                msg = (
                    "【工單分頁預覽與排版決策 (NX 原生備援視窗)】\n\n"
                    f"目前工段刀具數已超過 15 格：\n"
                    f"{stage_desc_str}\n\n"
                    f"預計總頁數對比：\n"
                    f"  - 單頁延伸模式：共 {total_extend_pages} 頁 (刀具列向下增加格數)\n"
                    f"  - 自動分頁模式：共 {total_paginate_pages} 頁 (標準 10 格/頁分頁)\n\n"
                    "請選擇排版模式：\n"
                    "・點選【是 (Yes)】：自動分頁顯示 (共 " + str(total_paginate_pages) + " 頁)\n"
                    "・點選【否 (No)】：單頁延伸顯示 (共 " + str(total_extend_pages) + " 頁)"
                )
                resp = the_ui.NXMessageBox.Show(
                    "工單分頁模式決策",
                    NXOpen.NXMessageBox.DialogType.Question,
                    msg
                )
                if resp == 1:
                    mode_result = "paginate"
                else:
                    mode_result = "extend"
        except Exception:
            mode_result = "extend"
    finally:
        if proc is not None:
            if proc.poll() is None:
                try:
                    proc.kill()
                    proc.wait(timeout=2)
                except Exception:
                    pass
            try:
                if proc.stdout: proc.stdout.close()
                if proc.stderr: proc.stderr.close()
                if proc.stdin: proc.stdin.close()
            except Exception:
                pass
        for tmp_f in [cfg_file, res_file]:
            if os.path.exists(tmp_f):
                try:
                    os.remove(tmp_f)
                except Exception:
                    pass

    return mode_result

def execute_nc_postprocessing(cam_setup, nc_tasks, post_config, listing=None):
    """
    依據使用者設定，逐一對 nc_tasks 執行 NX CAM 原生後處理轉出 NC 碼
    """
    if not nc_tasks or not cam_setup:
        return

    post_name = post_config.get("postprocessor_name", "").strip()
    custom_post = post_config.get("custom_post_path", "").strip()
    output_dir = post_config.get("output_dir", "").strip()
    extension = post_config.get("extension", ".nc").strip()

    if not extension.startswith("."):
        extension = "." + extension

    # 若為自訂後處理器檔案路徑
    actual_post = custom_post if (custom_post and os.path.exists(custom_post)) else post_name
    if actual_post.startswith("[自訂]"):
        actual_post = actual_post.replace("[自訂]", "").strip()

    if not output_dir:
        output_dir = os.getcwd()
    os.makedirs(output_dir, exist_ok=True)

    # 決定後處理單位 (PostDefined 優先，回退 Metric)
    units = getattr(NXOpen.CAM.CAMSetup.OutputUnits, "PostDefined", None)
    if units is None:
        units = getattr(NXOpen.CAM.CAMSetup.OutputUnits, "Metric", 0)

    if listing:
        listing.WriteLine("----------------------------------------")
        listing.WriteLine("【開始轉出 NC 碼】")
        listing.WriteLine(f"  - 後處理器：{actual_post}")
        listing.WriteLine(f"  - 輸出目錄：{output_dir}")
        listing.WriteLine(f"  - 程式副檔名：{extension}")
        listing.WriteLine(f"  - 待處理檔案數：{len(nc_tasks)} 個")

    success_count = 0
    fail_count = 0

    for idx, task in enumerate(nc_tasks, start=1):
        p_name = task.get("program_name", f"PROG_{idx}")
        ops = task.get("operations", [])
        out_file = os.path.join(output_dir, f"{p_name}{extension}")

        if not ops:
            if listing:
                listing.WriteLine(f"  [{idx:02d}/{len(nc_tasks):02d}] 略過 {p_name} (無有效工序)")
            continue

        try:
            if listing:
                listing.WriteLine(f"  [{idx:02d}/{len(nc_tasks):02d}] 正在後處理 {p_name}{extension} (共 {len(ops)} 道工序)...")

            # 調用 NX Open CAM 原生後處理 API
            cam_setup.Postprocess(ops, actual_post, out_file, units)

            if os.path.exists(out_file) and os.path.getsize(out_file) > 0:
                success_count += 1
                if listing:
                    listing.WriteLine(f"      ✔ 成功轉出：{out_file} ({os.path.getsize(out_file)} 位元組)")
            else:
                success_count += 1
                if listing:
                    listing.WriteLine(f"      ✔ 完成調用：{out_file}")
        except Exception as ex:
            fail_count += 1
            if listing:
                listing.WriteLine(f"      ✖ 後處理失敗 [{p_name}]：{str(ex)}")

    if listing:
        listing.WriteLine(f"NC 碼轉出完畢：成功 {success_count} 筆，失敗 {fail_count} 筆。")
        listing.WriteLine("----------------------------------------")

# ==================== 主執行流程 ====================

def main():
    the_session = NXOpen.Session.GetSession()
    the_ui = NXOpen.UI.GetUI()
    uf_session = NXOpen.UF.UFSession.GetUFSession()

    try:
        the_session.ListingWindow.Open()
    except Exception:
        pass

    # 安全取得工作零件 (支援 Work 與 Display 降級防護)
    work_part = None
    try:
        work_part = the_session.Parts.Work
    except NXOpen.NXException as ex:
        err_str = str(ex)
        if "使用者中止" in err_str or "abort" in err_str.lower():
            try:
                the_session.ListingWindow.WriteLine("提示：操作已由使用者中止。")
            except Exception:
                pass
            return
    except Exception:
        pass

    if work_part is None:
        try:
            work_part = the_session.Parts.Display
        except Exception:
            pass

    if work_part is None:
        try:
            the_session.ListingWindow.WriteLine("錯誤：未開啟任何工作零件，請確認零件處於開啟中！")
        except Exception:
            pass
        return

    full_part_path = work_part.FullPath
    if not full_part_path or not os.path.isabs(full_part_path):
        the_session.ListingWindow.WriteLine("錯誤：當前零件尚未存檔，請先儲存零件！")
        return

    sel_mgr = the_ui.SelectionManager
    num_selected = sel_mgr.GetNumSelectedObjects()

    if num_selected == 0:
        the_session.ListingWindow.WriteLine("提示：請先在 CAM 導覽器中選取要匯出的【程式群組/資料夾】或【工序】！")
        return

    target_folder = os.path.dirname(full_part_path)
    part_name = work_part.Leaf

    template_path = resolve_asset_file("ShopDoc_Template.xlsx", the_session=the_session, work_part=work_part)
    if not os.path.exists(template_path):
        the_session.ListingWindow.WriteLine(f"  [錯誤] 找不到 Excel 工單範本檔案：{template_path}，請確認已正確放置範本檔案！")
        return
    output_path = os.path.join(target_folder, f"{part_name}_選定工序工單.xlsx")

    the_session.ListingWindow.WriteLine("========================================")
    the_session.ListingWindow.WriteLine("開始分析 CAM 導覽器結構與工藝數據...")

    # 1. 遞迴收集並按群組分塊 (Chunks)，同時過濾排除群組與提取資訊群組
    selected_objects = []
    for i in range(num_selected):
        obj = sel_mgr.GetSelectedTaggedObject(i)
        if obj is not None:
            selected_objects.append(obj)

    # 嚴格保證選取物件之順序 100% 符合 CAM 導覽器面板由上至下 (Program Order) 之順序，並構建全域樹狀地圖
    cam_tree_map = {}
    extracted_info = {}
    try:
        cam_setup = work_part.CAMSetup
        if cam_setup:
            root_group = cam_setup.GetRoot(NXOpen.CAM.CAMSetup.View.ProgramOrder)
            if root_group:
                # 建立全域樹狀地圖 (精確鎖定每個節點所屬之頂層父資料夾與工段)，並自動全域掃描資訊資料夾 (ExcelTool 模式)
                cam_tree_map = build_cam_tree_map(root_group, extracted_info)

                order_map = {}
                order_counter = [0]
                def _build_order(node):
                    if not node:
                        return
                    t = getattr(node, "Tag", None)
                    if t and t not in order_map:
                        order_map[t] = order_counter[0]
                        order_counter[0] += 1
                    if isinstance(node, NXOpen.CAM.NCGroup):
                        try:
                            for m in node.GetMembers():
                                _build_order(m)
                        except Exception:
                            pass
                _build_order(root_group)
                selected_objects.sort(key=lambda o: order_map.get(getattr(o, "Tag", 0), 999999))
    except Exception:
        pass

    # 識別使用者直接點選之工序 (直接選取的工序，程式檔名將顯示工序自身名稱)
    directly_selected_tags = {obj.Tag for obj in selected_objects if isinstance(obj, NXOpen.CAM.Operation)}
    visited_tags = set()

    # 濾除已有祖先被選取之子物件，確保工單生成範圍與 NC 後處理轉出任務邊界 100% 對齊
    selected_set = set(selected_objects)
    def _has_selected_ancestor_node(node):
        curr = node
        while curr is not None:
            try:
                curr = curr.GetParent(NXOpen.CAM.CAMSetup.View.ProgramOrder)
            except Exception:
                curr = None
            if curr and curr in selected_set:
                return True
        return False

    top_selected_objects = [obj for obj in selected_objects if not _has_selected_ancestor_node(obj)]

    raw_chunks = []
    for tagged_obj in top_selected_objects:
        collect_cam_hierarchy(
            tagged_obj, raw_chunks, extracted_info,
            tree_map=cam_tree_map, visited_tags=visited_tags, directly_selected_tags=directly_selected_tags
        )

    if not raw_chunks:
        the_session.ListingWindow.WriteLine("錯誤：選取的項目中未包含任何有效的 CAM 工序！")
        return

    # 若有提取到資訊關鍵字，印出提示
    if extracted_info:
        the_session.ListingWindow.WriteLine(f"已自動識別表頭資訊：{extracted_info}")

    # 2. 建構 NC 轉出任務清單 (符合分組分檔、父資料夾包含子工序、單獨工序各自轉出、自動去重規則)
    nc_tasks = build_nc_tasks_from_selection(selected_objects)

    # 3. 彈出互動確認視窗：讓使用者決定是否依表單轉出 NC 碼、選擇機台後處理器與輸出路徑
    the_session.ListingWindow.WriteLine("----------------------------------------")
    the_session.ListingWindow.WriteLine("正在啟動後處理轉出與工單設定視窗...")
    post_config = invoke_nc_post_dialog(
        nc_tasks, target_folder, default_post="Fanuc_2026", listing=the_session.ListingWindow,
        the_session=the_session, work_part=work_part
    )

    action = post_config.get("action", "export_only")
    if action == "cancel":
        the_session.ListingWindow.WriteLine("\n[提示] 使用者已取消操作，流程中止。")
        return
    elif action == "post_and_export":
        the_session.ListingWindow.WriteLine("使用者選擇【轉出 NC 碼並匯出工單】。")
        try:
            execute_nc_postprocessing(cam_setup, nc_tasks, post_config, listing=the_session.ListingWindow)
        except Exception as ex:
            the_session.ListingWindow.WriteLine(f"執行後處理轉出時發生例外：{str(ex)}")
    else:
        the_session.ListingWindow.WriteLine("使用者選擇【僅匯出工單 (不轉 NC 碼)】。")

    # 2. 逐一萃取工序資訊並執行群組內刀具合併 (各子資料夾獨立合併，不跨子資料夾合併！)
    processed_chunks = []
    total_op_count = 0
    total_merged_step_count = 0
    last_folder_boundary_seen = None

    for chunk in raw_chunks:
        grp_name = chunk["group_name"]
        parent_prog = chunk.get("parent_program") or grp_name
        chunk_stage = chunk.get("stage", "通用工段")
        raw_ops = []

        for op in chunk["operations"]:
            # 工序名稱
            raw_op_name = op.Name
            clean_op_name = clean_program_name(raw_op_name)

            # 判定是否為使用者直接點選之工序
            is_direct = getattr(op, "_is_direct_op", False) or (getattr(op, "Tag", None) in directly_selected_tags)

            # 切削時間 (純秒數與格式化字串)
            op_seconds = get_operation_time_seconds(op)
            op_time_str = format_seconds_to_hms(op_seconds)

            # 刀具物件與參數
            tool_obj, tool_tag = get_tool_from_operation(op, uf_session)
            tool_name, tool_number, tool_diameter, flute_length, holder_length = get_tool_parameters(
                uf_session, tool_obj, tool_tag
            )

            # 轉速與進給 (以 try-finally 保證 NX C++ Builder 記憶體 100% 釋放，防止 ugraf.exe 核心記憶體洩漏)
            rpm_feed_note = ""
            feeds_builder = None
            try:
                feeds_builder = op.CreateFeedsBuilder()
                if feeds_builder:
                    rpm = f"{feeds_builder.SpindleRpmBuilder.Value:.0f}"
                    feed = f"{feeds_builder.CutFeedrateBuilder.Value:.1f}"
                    rpm_feed_note = f"S:{rpm} F:{feed}"
            except Exception:
                pass
            finally:
                if feeds_builder is not None:
                    try:
                        feeds_builder.Destroy()
                    except Exception:
                        pass

            raw_ops.append({
                "clean_op_name": clean_op_name,
                "is_direct_op": is_direct,
                "raw_tool_name": tool_name,
                "tool_name": tool_name,
                "tool_number": tool_number,
                "tool_diameter": tool_diameter,
                "flute_length": flute_length,
                "holder_length": holder_length,
                "time_seconds": op_seconds,
                "time": op_time_str,
                "note": rpm_feed_note
            })
            total_op_count += 1

        # 執行子群組內部刀號與工步智慧合併 (各子資料夾獨立合併，直接點選工序絕不合併)
        merged_ops = merge_consecutive_tools(raw_ops, group_name=grp_name)

        # 程式檔名規範：
        # 1. 若該工步為「直接點選之工序」：程式檔名填寫該工序自身名稱，每道工序皆各自獨立成列顯示
        # 2. 若該工步為「資料夾展開」：以母工段/父資料夾為邊界，僅首道工步填寫父資料夾名稱，後續所有工步全部留空
        for mop in merged_ops:
            is_direct = mop.get("is_direct_op", False)
            if is_direct:
                mop["op_name"] = mop.get("clean_op_name", "")
            else:
                folder_boundary = (chunk_stage, parent_prog)
                if folder_boundary != last_folder_boundary_seen:
                    mop["op_name"] = clean_program_name(parent_prog)
                    last_folder_boundary_seen = folder_boundary
                else:
                    mop["op_name"] = ""

        total_merged_step_count += len(merged_ops)

        # 日誌輸出提示
        the_session.ListingWindow.WriteLine(
            f"子資料夾 [{grp_name}] (所屬母程式: {parent_prog})：原始 {len(raw_ops)} 道工序，合併後共 {len(merged_ops)} 個工步："
        )
        for mop in merged_ops:
            seq_num = mop["seq"]
            p_name = mop["op_name"] if mop["op_name"] else f"(同上: {clean_program_name(parent_prog)})"
            t_num = mop["tool_number"]
            t_spec = mop.get("tool_spec_display", format_tool_display(mop.get("raw_tool_name"), mop.get("tool_diameter")))
            flen = mop["flute_length"]
            hlen = mop["holder_length"]
            t_str = mop["time"]
            mop["tool_spec_display"] = t_spec

            the_session.ListingWindow.WriteLine(
                f"  - 序號: {seq_num:02d} | 檔名: {p_name} | 刀號: {t_num} | 規格: {t_spec} | 刃長: {flen} | 夾長: {hlen} | 時間: {t_str}"
            )

        if merged_ops:
            processed_chunks.append({
                "group_name": grp_name,
                "parent_program": parent_prog,
                "stage": chunk.get("stage", "通用工段"),
                "operations": merged_ops
            })

    the_session.ListingWindow.WriteLine("----------------------------------------")
    the_session.ListingWindow.WriteLine(
        f"共讀取到 {total_op_count} 道有效工序，刀具合併後共 {total_merged_step_count} 個工步，正在計算分頁排版..."
    )

    # 3. 預先分析各工段工步數量與各模式下之預計頁數 (供使用者預覽與決策)
    stage_stats, total_extend_pages, total_paginate_pages, has_choice = preview_pagination_plan(
        processed_chunks, rows_per_page=ROWS_PER_PAGE
    )

    pagination_mode = "extend"  # 預設模式：單頁延伸

    if has_choice:
        the_session.ListingWindow.WriteLine("----------------------------------------")
        the_session.ListingWindow.WriteLine("【工單分頁預覽】：")
        for s in stage_stats:
            the_session.ListingWindow.WriteLine(
                f"  - 工段【{s['stage']}】({s['count']}刀)：單頁延伸規劃 ➔ {s['extend_desc']} | 自動分頁規劃 ➔ {s['paginate_desc']}"
            )
        the_session.ListingWindow.WriteLine(
            f"預計工單總頁數對比：單頁延伸模式為【共 {total_extend_pages} 頁】，自動分頁模式為【共 {total_paginate_pages} 頁】。"
        )
        the_session.ListingWindow.WriteLine("正在啟動【工單分頁預覽與排版決策視窗】供使用者選擇...")

        user_choice = invoke_pagination_prompt_dialog(
            stage_stats, total_extend_pages, total_paginate_pages, listing=the_session.ListingWindow,
            the_session=the_session, work_part=work_part
        )

        if user_choice == "cancel":
            the_session.ListingWindow.WriteLine("\n[提示] 使用者已取消排版操作，流程中止。")
            return
        elif user_choice == "paginate":
            pagination_mode = "paginate"
            the_session.ListingWindow.WriteLine(
                f"使用者選擇【📋 自動分頁顯示】(依標準 10 格/頁分頁，預計總共 {total_paginate_pages} 頁)。"
            )
        else:
            pagination_mode = "extend"
            the_session.ListingWindow.WriteLine(
                f"使用者選擇【📑 單頁延伸顯示】(刀具向下增加格數，預計總共 {total_extend_pages} 頁)。"
            )
    else:
        the_session.ListingWindow.WriteLine(
            f"各工段刀具數皆在 10 格以內，預先計算工單總頁數為：共 {total_extend_pages} 頁。"
        )

    # 執行智慧分頁 (支援工段優先分頁與動態增加格數/單頁延伸/截圖下一頁模式)
    pages = paginate_operations(processed_chunks, rows_per_page=ROWS_PER_PAGE, pagination_mode=pagination_mode)
    the_session.ListingWindow.WriteLine(f"分頁排版完成：最終產出為 {len(pages)} 頁。")

    # 4. 分析所有群組歸納獨立工段清單
    distinct_stages = get_distinct_stages(processed_chunks)
    the_session.ListingWindow.WriteLine(f"識別到的加工工段清單：{distinct_stages}")

    # 5. 方案二：【互動引導式拍照精靈】
    the_session.ListingWindow.WriteLine("----------------------------------------")
    the_session.ListingWindow.WriteLine("啟動【互動引導式拍照精靈】...")
    the_session.ListingWindow.WriteLine("提示：請在置頂拍照小視窗引導下，在 NX 主視窗旋轉縮放工件至最佳視角後點選拍照。")

    stage_images = {}
    try:
        stage_images = run_interactive_capture_wizard(
            distinct_stages, target_folder, uf_session, the_ui=the_ui, listing=the_session.ListingWindow, work_part=work_part
        )
        if stage_images:
            the_session.ListingWindow.WriteLine(f"各工段加工示圖已擷取：{list(stage_images.keys())}")
        else:
            the_session.ListingWindow.WriteLine("未截取示圖或已略過，工單展示區將保持留白。")
    except Exception as ex:
        the_session.ListingWindow.WriteLine(f"拍照精靈啟動異常 (安全跳過)：{str(ex)}")

    # 6. 判定表頭資訊 (依使用者最新規則)
    # 圖名：應為檔案名稱
    drawing_name = part_name
    # 圖號：若當中有 14 碼編號則為圖號，若無則留空
    drawing_number = extract_14_digit_drawing_number(part_name)
    # 尺寸：應該由取得設定的素材大小決定，若無資訊則留空
    blank_size = determine_blank_size(work_part, uf_session, raw_chunks, extracted_info)

    header_info = {
        "drawing_name": drawing_name,
        "drawing_number": drawing_number,
        "blank_size": blank_size,
        "part_number": extracted_info.get("part_number", ""),
        "holes": extracted_info.get("holes", ""),
        "thickness": extracted_info.get("thickness", "")
    }

    the_session.ListingWindow.WriteLine("----------------------------------------")
    the_session.ListingWindow.WriteLine(
        f"表頭解析結果：[圖名] {drawing_name} | "
        f"[圖號] {drawing_number if drawing_number else '(無14碼，留空)'} | "
        f"[尺寸] {blank_size if blank_size else '(未設定素材大小，留空)'}"
    )

    # 7. 透過 VBS 多頁寫入 Excel 並內嵌加工示圖
    the_session.ListingWindow.WriteLine("正在產生多頁 Excel 工單與嵌入加工示圖...")
    if stage_images:
        the_session.ListingWindow.WriteLine(f"準備嵌入之加工示圖列表：{stage_images}")
    try:
        export_multipage_via_vbs(pages, template_path, output_path, work_part, header_info, stage_images=stage_images)
        the_session.ListingWindow.WriteLine(f"工單建立完成！檔案路徑：{output_path}")
        the_session.ListingWindow.WriteLine("========================================")
        os.startfile(output_path)
    except Exception as ex:
        the_session.ListingWindow.WriteLine(f"匯出失敗：{str(ex)}")
    finally:
        # Zero-Debug 資源防護與記憶體清理：
        # 1. 拍照精靈產生的暫存圖檔已實體內嵌於 Excel，安全刪除外部暫存檔保持目錄整潔
        for img_f in stage_images.values():
            if img_f and os.path.exists(img_f):
                try:
                    os.remove(img_f)
                except Exception:
                    pass

        # 2. 清理目標目錄中所有本次殘留之拍照暫存圖檔與擺正旗標
        try:
            for f in os.listdir(target_folder):
                if (f.startswith("_temp_stage_") and f.endswith(".png")) or (f.startswith("_snap_request_") and f.endswith(".flag")):
                    fp = os.path.join(target_folder, f)
                    try:
                        os.remove(fp)
                    except Exception:
                        pass
        except Exception:
            pass

        # 3. 顯式解構大型資料集合參照，釋放 NXOpen 封裝物件
        try:
            raw_chunks.clear()
            processed_chunks.clear()
            pages.clear()
            cam_tree_map.clear()
            selected_objects.clear()
        except Exception:
            pass

        # 4. 主動觸發 Python 垃圾回收 (Garbage Collection)，徹底釋放 NX 內部記憶體堆疊
        try:
            import gc
            gc.collect()
        except Exception:
            pass

        # 5. 安全關閉資訊視窗 (ListingWindow)，釋放 NX 內部關聯之 QtWebEngine 瀏覽器資源
        try:
            if the_session is not None and hasattr(the_session, "ListingWindow"):
                the_session.ListingWindow.Close()
        except Exception:
            pass

if __name__ == "__main__":
    try:
        main()
    except NXOpen.NXException as ex:
        err_msg = str(ex)
        try:
            the_session = NXOpen.Session.GetSession()
            the_session.ListingWindow.Open()
            if "使用者中止" in err_msg or "abort" in err_msg.lower():
                the_session.ListingWindow.WriteLine("\n[提示] 操作已中止。")
            else:
                the_session.ListingWindow.WriteLine(f"\n[NX 執行中斷] {err_msg}")
        except Exception:
            pass
    except Exception as ex:
        try:
            the_session = NXOpen.Session.GetSession()
            the_session.ListingWindow.Open()
            the_session.ListingWindow.WriteLine(f"\n[執行未預期異常] {str(ex)}")
        except Exception:
            pass