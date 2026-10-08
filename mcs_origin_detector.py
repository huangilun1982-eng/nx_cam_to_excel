# -*- coding: utf-8 -*-
"""
組件名稱：mcs_origin_detector.py
功能職責：NX CAM 加工座標系 (MCS / G54) 與原點方位自動解析引擎
設計原則：
1. 組件功能單純原則：專注於從工序回溯 OrientGeometry、讀取夾具偏置與幾何邊界盒，判定原點方位。
2. 方案 2 階梯防禦架構 (CAM 工件/素材優先，未指定時抓顯示物件)：
   - 階梯 1 (第一優先)：CAM WORKPIECE / MillGeom 所指定之 Part / Blank 實體幾何。
   - 階梯 2 (第二優先)：未指定 CAM 幾何時，抓圖面當前「顯示」物件 (Visible & Not Blanked)。
   - 階梯 3 (第三保底)：全零件所有實體 (All Bodies Fallback)。
   - 階梯 4 (第四保底)：無實體時回退預設標註 (加工原點【G54】：X=MID, Y=MID, Z=TOP)。
3. 零破壞與多階防禦 (Zero-Debug Protocol)：各層級操作皆設有防護機制，絕不中斷流程。
4. 輸出規範：嚴格遵循使用者指定格式：
   - X 方向：0, MID, MAX
   - Y 方向：0, MID, MAX
   - Z 方向：0, MID, TOP
   - 範例：加工原點【G54】：X=MID, Y=MID, Z=TOP
"""

import math

def parse_fixture_offset_to_gcode(val):
    """
    將 NX Fixture Offset 整數數值轉換為標準 CNC G 碼代號 (G54, G55...)
    """
    try:
        iv = int(val)
    except Exception:
        return "G54"

    if iv in [0, 1, 54]:
        return "G54"
    elif iv in [2, 55]:
        return "G55"
    elif iv in [3, 56]:
        return "G56"
    elif iv in [4, 57]:
        return "G57"
    elif iv in [5, 58]:
        return "G58"
    elif iv in [6, 59]:
        return "G59"
    elif iv >= 54:
        return f"G{iv}"
    elif iv > 6:
        return f"G54.1 P{iv - 6}"
    return "G54"

def determine_axis_orientation(origin_val, min_val, max_val, is_z_axis=False, tol=0.1):
    """
    比對原點座標與幾何邊界盒 (Bounding Box)，判定方位標籤。
    
    規則：
      X/Y 軸：0, MID, MAX
      Z 軸：0, MID, TOP
    """
    mid_val = (min_val + max_val) / 2.0

    if is_z_axis:
        # Z 軸判定 (0, MID, TOP)
        if abs(origin_val - max_val) <= tol:
            return "TOP"
        elif abs(origin_val - mid_val) <= tol:
            return "MID"
        elif abs(origin_val - min_val) <= tol or abs(origin_val - 0.0) <= tol:
            return "0"
        elif origin_val >= max_val:
            return "TOP"
        elif origin_val <= min_val:
            return "0"
        else:
            return "MID"
    else:
        # X / Y 軸判定 (0, MID, MAX)
        if abs(origin_val - mid_val) <= tol:
            return "MID"
        elif abs(origin_val - min_val) <= tol or (abs(origin_val - 0.0) <= tol and abs(min_val - 0.0) <= tol):
            return "0"
        elif abs(origin_val - max_val) <= tol:
            return "MAX"
        elif abs(origin_val - 0.0) <= tol:
            return "0"
        else:
            dist_min = abs(origin_val - min_val)
            dist_mid = abs(origin_val - mid_val)
            dist_max = abs(origin_val - max_val)
            if dist_mid <= dist_min and dist_mid <= dist_max:
                return "MID"
            elif dist_min <= dist_max:
                return "0"
            else:
                return "MAX"

def extract_bounding_box_for_entities(entities, uf_session):
    """
    計算一組實體或幾何物件 (Body / Face / Tag) 的整體聯集邊界盒。
    
    回傳：
      (found, min_coords, max_coords)
    """
    if not entities or not uf_session:
        return False, None, None

    all_min = [1e9, 1e9, 1e9]
    all_max = [-1e9, -1e9, -1e9]
    found = False

    for ent in entities:
        tag = getattr(ent, "Tag", ent) if not isinstance(ent, int) else ent
        try:
            b_res = uf_session.ModlGeneral.AskBoundingBox(tag)
            if b_res and len(b_res) == 6:
                found = True
                for i in range(3):
                    if b_res[i] < all_min[i]:
                        all_min[i] = b_res[i]
                    if b_res[i+3] > all_max[i]:
                        all_max[i] = b_res[i+3]
        except Exception:
            pass

    if found:
        return True, all_min, all_max
    return False, None, None

def find_associated_workpiece_node(first_op_obj, cam_setup):
    """
    尋找與工序相關聯的 WORKPIECE / MillGeom 節點。
    1. 優先由工序向上回溯 Geometry 父節點。
    2. 次選遍歷幾何樹尋找 WORKPIECE 節點。
    """
    if not cam_setup:
        return None

    # 1. 工序向上回溯
    if first_op_obj is not None:
        curr = first_op_obj
        depth = 0
        while curr and depth < 10:
            depth += 1
            try:
                import NXOpen.CAM
                if isinstance(curr, NXOpen.CAM.Operation):
                    parent = curr.GetParent(NXOpen.CAM.CAMSetup.View.Geometry)
                else:
                    parent = curr.GetParent()

                if not parent:
                    break
                p_name = str(getattr(parent, "Name", "")).upper()
                if "WORKPIECE" in p_name:
                    return parent
                curr = parent
            except Exception:
                break

    # 2. Geometry Root 全域遍歷搜尋
    try:
        try:
            import NXOpen.CAM
            view_geom = NXOpen.CAM.CAMSetup.View.Geometry
            geom_root = cam_setup.GetRoot(view_geom)
        except Exception:
            geom_root = cam_setup.GetRoot() if hasattr(cam_setup, "GetRoot") else None

        if geom_root:
            def scan_node(node):
                if "WORKPIECE" in str(getattr(node, "Name", "")).upper():
                    return node
                try:
                    for m in node.GetMembers():
                        found = scan_node(m)
                        if found:
                            return found
                except Exception:
                    pass
                return None
            return scan_node(geom_root)
    except Exception:
        pass

    return None

def extract_target_bounding_box(work_part, cam_setup, uf_session, first_op_obj=None, mcs_node=None):
    """
    方案 2 核心：幾何邊界盒階梯提取引擎
    階梯 1 (第一優先)：CAM WORKPIECE 指定之 Part / Blank 幾何實體
    階梯 2 (第二優先)：圖面當前顯示物件 (可見且未隱藏之實體)
    階梯 3 (第三保底)：全零件所有實體
    
    回傳：
      (found, min_coords, max_coords, source_desc)
    """
    # -------------------------------------------------------------
    # 階梯 1：CAM 工件 / 素材指定實體 (CAM Geometry Selection)
    # -------------------------------------------------------------
    try:
        wp_node = find_associated_workpiece_node(first_op_obj, cam_setup)
        if wp_node and cam_setup:
            builder = cam_setup.CAMGroupCollection.CreateMillGeomBuilder(wp_node)
            cam_entities = []
            try:
                geoms_to_check = []
                if hasattr(builder, "BlankGeometry"):
                    geoms_to_check.append(builder.BlankGeometry)
                if hasattr(builder, "PartGeometry"):
                    geoms_to_check.append(builder.PartGeometry)

                for g in geoms_to_check:
                    try:
                        glist = g.GeometryList
                        length = glist.Length if hasattr(glist, "Length") else (glist.Count if hasattr(glist, "Count") else len(glist))
                        for i in range(length):
                            gset = glist.FindItem(i)
                            items = []
                            try:
                                items = list(gset.GetItems())
                            except Exception:
                                pass
                            if not items and hasattr(gset, "ScCollector"):
                                try:
                                    items = list(gset.ScCollector.GetObjects())
                                except Exception:
                                    pass
                            for it in items:
                                if it not in cam_entities:
                                    cam_entities.append(it)
                    except Exception:
                        pass
            finally:
                try:
                    builder.Destroy()
                except Exception:
                    pass

            if cam_entities:
                found, b_min, b_max = extract_bounding_box_for_entities(cam_entities, uf_session)
                if found:
                    return True, b_min, b_max, "CAM_GEOMETRY"
    except Exception:
        pass

    # -------------------------------------------------------------
    # 階梯 2：圖面當前顯示實體 (Visible Bodies in Current Display)
    # 條件：所屬圖層不是隱藏 (Layer Status != 4) 且 物件非 Blanked
    # -------------------------------------------------------------
    if work_part and hasattr(work_part, "Bodies"):
        try:
            visible_bodies = []
            for b in work_part.Bodies:
                try:
                    # 圖層狀態檢查：4 代表隱藏/不可見 (UF_LAYER_INACTIVE / INVIS)
                    layer_st = 1
                    if uf_session and hasattr(uf_session, "Layer"):
                        b_layer = getattr(b, "Layer", 1)
                        layer_st = uf_session.Layer.AskStatus(b_layer)
                    
                    is_blanked = getattr(b, "IsBlanked", False)
                    if layer_st != 4 and not is_blanked:
                        visible_bodies.append(b)
                except Exception:
                    visible_bodies.append(b)

            if visible_bodies:
                found, b_min, b_max = extract_bounding_box_for_entities(visible_bodies, uf_session)
                if found:
                    return True, b_min, b_max, "DISPLAY_BODIES"
        except Exception:
            pass

    # -------------------------------------------------------------
    # 階梯 3：全零件所有實體保底 (All Bodies Fallback)
    # -------------------------------------------------------------
    if work_part and hasattr(work_part, "Bodies"):
        try:
            all_bodies = list(work_part.Bodies)
            if all_bodies:
                found, b_min, b_max = extract_bounding_box_for_entities(all_bodies, uf_session)
                if found:
                    return True, b_min, b_max, "ALL_BODIES"
        except Exception:
            pass

    # -------------------------------------------------------------
    # 階梯 4：無任何實體
    # -------------------------------------------------------------
    return False, None, None, "NONE"

def resolve_stage_mcs_origin_string(raw_ops, work_part=None, cam_setup=None, uf_session=None):
    """
    解析工段對應之加工原點說明字串。
    
    參數：
      raw_ops: 該工段包含的工序清單 (包含 NX CAM Operation 物件，若有)
      work_part: 當前工作零件 (NXOpen.Part)
      cam_setup: CAM 設置 (NXOpen.CAM.CAMSetup)
      uf_session: NX UFSession
      
    回傳：
      origin_str: 例如 "加工原點【G54】：X=MID, Y=MID, Z=TOP"
    """
    fallback_result = "加工原點【G54】：X=MID, Y=MID, Z=TOP"

    if not work_part or not cam_setup or not uf_session:
        return fallback_result

    # 1. 從該工段工步中尋找有效的 NX CAM Operation 物件
    first_op_obj = None
    if raw_ops:
        for op_item in raw_ops:
            op_candidate = op_item.get("_nx_op_obj") if isinstance(op_item, dict) else op_item
            if op_candidate is not None:
                first_op_obj = op_candidate
                break

    # 2. 向上回溯幾何父節點尋找 OrientGeometry (MCS 節點)
    mcs_node = None
    if first_op_obj is not None:
        curr = first_op_obj
        depth = 0
        while curr and depth < 10:
            depth += 1
            try:
                import NXOpen.CAM
                if isinstance(curr, NXOpen.CAM.Operation):
                    parent = curr.GetParent(NXOpen.CAM.CAMSetup.View.Geometry)
                else:
                    parent = curr.GetParent()

                if not parent:
                    break
                if isinstance(parent, NXOpen.CAM.OrientGeometry):
                    mcs_node = parent
                    break
                curr = parent
            except Exception:
                break

    # 備援：若無工步回溯，自 Geometry Root 直接搜尋第一個 OrientGeometry
    if not mcs_node:
        try:
            import NXOpen.CAM
            geom_root = cam_setup.GetRoot(NXOpen.CAM.CAMSetup.View.Geometry)
            if geom_root:
                def scan_mcs(node):
                    if isinstance(node, NXOpen.CAM.OrientGeometry):
                        return node
                    try:
                        for m in node.GetMembers():
                            found = scan_mcs(m)
                            if found:
                                return found
                    except Exception:
                        pass
                    return None
                mcs_node = scan_mcs(geom_root)
        except Exception:
            pass

    if not mcs_node:
        return fallback_result

    # 3. 讀取 MCS 資訊 (原點向量與夾具偏置值)
    g_code = "G54"
    origin_x = 0.0
    origin_y = 0.0
    origin_z = 0.0
    has_origin = False

    try:
        # 確保 CAM Session 初始化
        try:
            uf_session.Cam.InitSession()
        except Exception:
            pass

        builder = cam_setup.CAMGroupCollection.CreateMillOrientGeomBuilder(mcs_node)
        if builder:
            try:
                if hasattr(builder, "FixtureOffsetBuilder"):
                    fob_val = builder.FixtureOffsetBuilder.Value
                    g_code = parse_fixture_offset_to_gcode(fob_val)
                if hasattr(builder, "Mcs") and builder.Mcs:
                    origin_x = builder.Mcs.Origin.X
                    origin_y = builder.Mcs.Origin.Y
                    origin_z = builder.Mcs.Origin.Z
                    has_origin = True
            finally:
                try:
                    builder.Destroy()
                except Exception:
                    pass
    except Exception:
        pass

    if not has_origin:
        return f"加工原點【{g_code}】：X=MID, Y=MID, Z=TOP"

    # 4. 依照方案 2 階梯引擎提取目標邊界盒 (CAM 優先 -> 顯示物件 -> 所有實體)
    found_box, b_min, b_max, source = extract_target_bounding_box(
        work_part=work_part,
        cam_setup=cam_setup,
        uf_session=uf_session,
        first_op_obj=first_op_obj,
        mcs_node=mcs_node
    )

    if not found_box or not b_min or not b_max:
        # 若無法讀取邊界盒，預設四面分中頂面碰刀
        return f"加工原點【{g_code}】：X=MID, Y=MID, Z=TOP"

    box_x_min, box_y_min, box_z_min = b_min[0], b_min[1], b_min[2]
    box_x_max, box_y_max, box_z_max = b_max[0], b_max[1], b_max[2]

    # 5. 幾何比對運算
    x_orient = determine_axis_orientation(origin_x, box_x_min, box_x_max, is_z_axis=False)
    y_orient = determine_axis_orientation(origin_y, box_y_min, box_y_max, is_z_axis=False)
    z_orient = determine_axis_orientation(origin_z, box_z_min, box_z_max, is_z_axis=True)

    return f"加工原點【{g_code}】：X={x_orient}, Y={y_orient}, Z={z_orient}"
