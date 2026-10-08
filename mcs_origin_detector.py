# -*- coding: utf-8 -*-
"""
組件名稱：mcs_origin_detector.py
功能職責：NX CAM 加工座標系 (MCS / G54) 與原點方位自動解析引擎
設計原則：
1. 100% 基於「加工座標系 (MCS)」幾何空間轉換：
   - 徹底杜絕將世界絕對座標 (ACS) 誤當加工座標比對之錯誤。
   - 提取 MCS 之原點向量與三軸姿態向量 (X軸, Y軸, Z軸/主軸刀具進給方向)。
   - 將工件/素材實體幾何投影轉換至 MCS 局部空間，求得實體在加工座標系下的局部邊界盒 [X', Y', Z']。
   - 加工原點在 MCS 局部空間座標為 (0, 0, 0)，以 0.0 為基準比對實體局部邊界，精準判定 Z=TOP / 0 / MID 及 X/Y 方位。
2. 方案 2 階梯防禦架構 (CAM 工件/素材優先，未指定時抓顯示物件)：
   - 階梯 1 (第一優先)：CAM WORKPIECE / MillGeom 所指定之 Part / Blank 實體幾何。
   - 階梯 2 (第二優先)：未指定 CAM 幾何時，抓圖面當前「顯示」物件 (Visible & Not Blanked)。
   - 階梯 3 (第三保底)：全零件所有實體 (All Bodies Fallback)。
   - 階梯 4 (第四保底)：無實體時回退預設標註 (加工原點【G54】：X=MID, Y=MID, Z=TOP)。
3. 零破壞與多階防禦 (Zero-Debug Protocol)：各層級操作皆設有防護機制，絕不中斷流程。
4. 輸出規範：嚴格遵循標準格式：
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


def determine_axis_orientation(origin_val, min_val, max_val, is_z_axis=False, tol=0.5):
    """
    比對原點座標與幾何邊界盒 (Bounding Box)，依據相對空間位置與特徵距離判定方位標籤。
    
    規則：
      X/Y 軸：0 (最小側邊緣), MID (中心分中), MAX (最大側邊緣)
      Z 軸：0 (底面碰刀/台面基準), MID (厚度分中), TOP (頂面碰刀/上表面基準)
      
    設計重點：
      1. 100% 依據原點相對於目標外包盒的空間位置決定，預設 0.5mm 現場加工公差容錯。
      2. 支援原點高於頂面 (預留切削量)、低於底面 (台面下偏置) 的邊界吸附。
      3. 在實體厚度內部時，以最近特徵距離 (Nearest Distance) 判定歸屬區域。
    """
    # 安全防護：若邊界盒無效或顛倒，自動修正
    if min_val > max_val:
        min_val, max_val = max_val, min_val

    mid_val = (min_val + max_val) / 2.0

    if is_z_axis:
        # Z 軸判定 (0, MID, TOP)
        # 1. 頂面邊界與上方區域吸附 (原點大於等於頂面減去容錯值，包含預留切削量)
        if origin_val >= (max_val - tol):
            return "TOP"
        # 2. 底面邊界與下方區域吸附 (原點小於等於底面加上容錯值，包含台面碰刀)
        if origin_val <= (min_val + tol):
            return "0"
        # 3. 中心位置吸附 (中心公差內)
        if abs(origin_val - mid_val) <= tol:
            return "MID"
        
        # 4. 厚度內部依最近距離判定歸屬區域
        dist_top = abs(origin_val - max_val)
        dist_mid = abs(origin_val - mid_val)
        dist_bot = abs(origin_val - min_val)

        if dist_top <= dist_mid and dist_top <= dist_bot:
            return "TOP"
        elif dist_mid <= dist_top and dist_mid <= dist_bot:
            return "MID"
        else:
            return "0"
    else:
        # X / Y 軸判定 (0, MID, MAX)
        # 1. 中心分中吸附
        if abs(origin_val - mid_val) <= tol:
            return "MID"
        # 2. 最小側邊緣吸附
        if origin_val <= (min_val + tol):
            return "0"
        # 3. 最大側邊緣吸附
        if origin_val >= (max_val - tol):
            return "MAX"
        
        # 4. 內部依最近距離判定
        dist_min = abs(origin_val - min_val)
        dist_mid = abs(origin_val - mid_val)
        dist_max = abs(origin_val - max_val)

        if dist_mid <= dist_min and dist_mid <= dist_max:
            return "MID"
        elif dist_min <= dist_max:
            return "0"
        else:
            return "MAX"


def project_point_to_csys(pt, origin, x_vec, y_vec, z_vec):
    """
    將三維絕對空間坐標點 pt 投影至以 origin 為原點、(x_vec, y_vec, z_vec) 為正交軸向量之局部坐標系中。
    回傳 (x', y', z')
    """
    dx = pt[0] - origin[0]
    dy = pt[1] - origin[1]
    dz = pt[2] - origin[2]

    px = dx * x_vec[0] + dy * x_vec[1] + dz * x_vec[2]
    py = dx * y_vec[0] + dy * y_vec[1] + dz * y_vec[2]
    pz = dx * z_vec[0] + dy * z_vec[1] + dz * z_vec[2]
    return px, py, pz


def get_mcs_geometry_axes(builder, mcs_node, uf_session=None):
    """
    精確提取加工座標系 (MCS) 之原點與姿態向量。
    回傳:
      (origin, x_vec, y_vec, z_vec, csys_tag)
      origin: [Ox, Oy, Oz]
      x_vec: [Xx, Xy, Xz] (單位向量)
      y_vec: [Yx, Yy, Yz] (單位向量)
      z_vec: [Zx, Zy, Zz] (單位向量，主軸刀具進給方向)
    """
    origin = [0.0, 0.0, 0.0]
    x_vec = [1.0, 0.0, 0.0]
    y_vec = [0.0, 1.0, 0.0]
    z_vec = [0.0, 0.0, 1.0]
    csys_tag = None

    # 1. 優先由 OrientGeomBuilder 的 Mcs 物件讀取
    try:
        if builder and hasattr(builder, "Mcs") and builder.Mcs:
            mcs_obj = builder.Mcs
            csys_tag = getattr(mcs_obj, "Tag", None)
            
            # 讀取原點
            if hasattr(mcs_obj, "Origin"):
                origin = [float(mcs_obj.Origin.X), float(mcs_obj.Origin.Y), float(mcs_obj.Origin.Z)]
                
            # 讀取姿態 (Matrix3x3)
            if hasattr(mcs_obj, "Orientation"):
                orient = mcs_obj.Orientation
                if hasattr(orient, "Xx"):
                    x_vec = [float(orient.Xx), float(orient.Xy), float(orient.Xz)]
                    y_vec = [float(orient.Yx), float(orient.Yy), float(orient.Yz)]
                    z_vec = [float(orient.Zx), float(orient.Zy), float(orient.Zz)]
    except Exception:
        pass

    # 2. 備援：透過 UFSession.Csys 讀取
    if csys_tag and uf_session and hasattr(uf_session, "Csys"):
        try:
            mtx_tag, orig_arr = uf_session.Csys.AskCsysInfo(csys_tag)
            if orig_arr and len(orig_arr) >= 3:
                origin = [float(orig_arr[0]), float(orig_arr[1]), float(orig_arr[2])]
            if mtx_tag:
                vals = uf_session.Csys.AskMatrixValues(mtx_tag)
                if vals and len(vals) >= 9:
                    x_vec = [float(vals[0]), float(vals[1]), float(vals[2])]
                    y_vec = [float(vals[3]), float(vals[4]), float(vals[5])]
                    z_vec = [float(vals[6]), float(vals[7]), float(vals[8])]
        except Exception:
            pass

    # 規範化向量長度防護 (確保為單位向量)
    def normalize(v, default):
        mag = math.sqrt(v[0]*v[0] + v[1]*v[1] + v[2]*v[2])
        if mag > 1e-6:
            return [v[0]/mag, v[1]/mag, v[2]/mag]
        return default

    x_vec = normalize(x_vec, [1.0, 0.0, 0.0])
    y_vec = normalize(y_vec, [0.0, 1.0, 0.0])
    z_vec = normalize(z_vec, [0.0, 0.0, 1.0])

    return origin, x_vec, y_vec, z_vec, csys_tag


def extract_mcs_bounding_box_for_entities(entities, uf_session, origin, x_vec, y_vec, z_vec, csys_tag=None):
    """
    計算一組實體或幾何物件在加工座標系 (MCS) 局部空間中的邊界盒 [X', Y', Z']。
    
    回傳：
      (found, min_local, max_local)
      min_local: [min_x', min_y', min_z']
      max_local: [max_x', max_y', max_z']
    """
    if not entities:
        return False, None, None

    all_min = [1e9, 1e9, 1e9]
    all_max = [-1e9, -1e9, -1e9]
    found = False

    for ent in entities:
        tag = getattr(ent, "Tag", ent) if not isinstance(ent, int) else ent
        ent_found = False

        # --- 策略 A：嘗試 NX 原廠 AskBoundingBoxExact API ---
        if csys_tag and uf_session and hasattr(uf_session, "Modl") and hasattr(uf_session.Modl, "AskBoundingBoxExact"):
            try:
                res = uf_session.Modl.AskBoundingBoxExact(tag, csys_tag)
                # res 格式通常為 (min_corner, directions, distances)
                if res and len(res) >= 3:
                    min_c, dirs, dists = res[0], res[1], res[2]
                    # 若已直接取得局部尺寸
                    if len(min_c) >= 3 and len(dists) >= 3:
                        c_min = [float(min_c[0]), float(min_c[1]), float(min_c[2])]
                        c_max = [c_min[0] + float(dists[0]), c_min[1] + float(dists[1]), c_min[2] + float(dists[2])]
                        for i in range(3):
                            if c_min[i] < all_min[i]: all_min[i] = c_min[i]
                            if c_max[i] > all_max[i]: all_max[i] = c_max[i]
                        found = True
                        ent_found = True
            except Exception:
                pass

        if ent_found:
            continue

        # --- 策略 B：提取實體幾何頂點投影至 MCS 空間 ---
        body_points = []
        try:
            edges = []
            if hasattr(ent, "GetEdges"):
                edges = ent.GetEdges()
            elif hasattr(ent, "Edges"):
                edges = ent.Edges
            
            for e in edges:
                try:
                    if hasattr(e, "GetVertices"):
                        v1, v2 = e.GetVertices()
                        if v1: body_points.append([float(v1.X), float(v1.Y), float(v1.Z)])
                        if v2: body_points.append([float(v2.X), float(v2.Y), float(v2.Z)])
                except Exception:
                    pass
        except Exception:
            pass

        # --- 策略 C：若無頂點或為特徵幾何，以世界邊界盒 8 個角點投影 ---
        if not body_points and uf_session and hasattr(uf_session, "ModlGeneral"):
            try:
                b_res = uf_session.ModlGeneral.AskBoundingBox(tag)
                if b_res and len(b_res) == 6:
                    xs = [b_res[0], b_res[3]]
                    ys = [b_res[1], b_res[4]]
                    zs = [b_res[2], b_res[5]]
                    for bx in xs:
                        for by in ys:
                            for bz in zs:
                                body_points.append([bx, by, bz])
            except Exception:
                pass

        # 執行投影轉換
        if body_points:
            for pt in body_points:
                px, py, pz = project_point_to_csys(pt, origin, x_vec, y_vec, z_vec)
                found = True
                if px < all_min[0]: all_min[0] = px
                if px > all_max[0]: all_max[0] = px
                if py < all_min[1]: all_min[1] = py
                if py > all_max[1]: all_max[1] = py
                if pz < all_min[2]: all_min[2] = pz
                if pz > all_max[2]: all_max[2] = pz

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
                view_geom = NXOpen.CAM.CAMSetup.View.Geometry
                parent = None
                try:
                    parent = curr.GetParent(view_geom)
                except Exception:
                    try:
                        parent = curr.GetParent()
                    except Exception:
                        parent = None

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
        import NXOpen.CAM
        view_geom = NXOpen.CAM.CAMSetup.View.Geometry
        geom_root = cam_setup.GetRoot(view_geom)
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


def extract_target_bounding_box_in_mcs(work_part, cam_setup, uf_session, origin, x_vec, y_vec, z_vec, csys_tag=None, first_op_obj=None, mcs_node=None):
    """
    方案 2 核心：幾何實體投影至加工座標系 (MCS) 階梯提取引擎
    階梯 1 (第一優先)：CAM WORKPIECE 指定之 Part / Blank 幾何實體
    階梯 2 (第二優先)：圖面當前顯示實體 (可見且未隱藏之實體)
    階梯 3 (第三保底)：全零件所有實體
    
    回傳：
      (found, min_local, max_local, source_desc)
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
                found, b_min, b_max = extract_mcs_bounding_box_for_entities(
                    cam_entities, uf_session, origin, x_vec, y_vec, z_vec, csys_tag
                )
                if found:
                    return True, b_min, b_max, "CAM_GEOMETRY"
    except Exception:
        pass

    # -------------------------------------------------------------
    # 階梯 2：圖面當前顯示實體 (Visible Bodies in Current Display)
    # 條件：物件非 Blanked 且 所屬圖層為合法工作圖層 (1~256 且 Layer Status != 4)
    # -------------------------------------------------------------
    if work_part and hasattr(work_part, "Bodies"):
        try:
            visible_bodies = []
            for b in work_part.Bodies:
                try:
                    if getattr(b, "IsBlanked", False):
                        continue

                    b_layer = getattr(b, "Layer", 1)
                    if not (1 <= b_layer <= 256):
                        continue

                    if uf_session and hasattr(uf_session, "Layer"):
                        layer_st = uf_session.Layer.AskStatus(b_layer)
                        if layer_st == 4:
                            continue

                    visible_bodies.append(b)
                except Exception:
                    continue

            if visible_bodies:
                found, b_min, b_max = extract_mcs_bounding_box_for_entities(
                    visible_bodies, uf_session, origin, x_vec, y_vec, z_vec, csys_tag
                )
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
                found, b_min, b_max = extract_mcs_bounding_box_for_entities(
                    all_bodies, uf_session, origin, x_vec, y_vec, z_vec, csys_tag
                )
                if found:
                    return True, b_min, b_max, "ALL_BODIES"
        except Exception:
            pass

    # -------------------------------------------------------------
    # 階梯 4：無任何實體
    # -------------------------------------------------------------
    return False, None, None, "NONE"


def extract_target_bounding_box(work_part, cam_setup, uf_session, origin=None, x_vec=None, y_vec=None, z_vec=None, csys_tag=None, first_op_obj=None, mcs_node=None):
    """
    通用相容介面：若未指定 MCS 參數，則預設以世界原點與單位座標軸執行邊界盒計算。
    """
    if origin is None:
        origin = [0.0, 0.0, 0.0]
    if x_vec is None:
        x_vec = [1.0, 0.0, 0.0]
    if y_vec is None:
        y_vec = [0.0, 1.0, 0.0]
    if z_vec is None:
        z_vec = [0.0, 0.0, 1.0]

    return extract_target_bounding_box_in_mcs(
        work_part=work_part,
        cam_setup=cam_setup,
        uf_session=uf_session,
        origin=origin,
        x_vec=x_vec,
        y_vec=y_vec,
        z_vec=z_vec,
        csys_tag=csys_tag,
        first_op_obj=first_op_obj,
        mcs_node=mcs_node
    )


def is_mcs_geometry_node(node):
    """
    判定一個 CAM 幾何節點是否為加工座標系節點 (OrientGeometry / MCS)。
    支援原生型別判定、名稱識別 (MCS_*) 與子型別相容。
    """
    if node is None:
        return False
    try:
        import NXOpen.CAM
        if isinstance(node, NXOpen.CAM.OrientGeometry):
            return True
    except Exception:
        pass
    type_name = type(node).__name__.upper()
    if "ORIENT" in type_name or "MCS" in type_name:
        return True
    node_name = str(getattr(node, "Name", "")).upper()
    if node_name.startswith("MCS") or "_MCS" in node_name or "MCS_" in node_name:
        return True
    return False


def resolve_stage_mcs_origin_string(raw_ops, work_part=None, cam_setup=None, uf_session=None):
    """
    解析工段對應之加工原點說明字串 (100% 基於加工座標系 MCS 幾何轉換)。
    
    參數：
      raw_ops: 該工段包含的工序清單 (包含 NX CAM Operation 物件，若有)
      work_part: 當前工作零件 (NXOpen.Part)
      cam_setup: CAM 設置 (NXOpen.CAM.CAMSetup)
      uf_session: NX UFSession
      
    回傳：
      origin_str: 例如 "加工原點【G54】：X=MID, Y=MID, Z=TOP"
    """
    fallback_result = "加工原點【G54】：X=MID, Y=MID, Z=TOP"

    if not work_part or not cam_setup:
        return fallback_result

    # 1. 向上回溯幾何父節點尋找該工段專屬之 OrientGeometry (MCS 節點)
    mcs_node = None
    first_op_obj = None
    for op_item in (raw_ops or []):
        op_obj = op_item.get("_nx_op_obj") if isinstance(op_item, dict) else op_item
        if op_obj is None:
            continue
        if first_op_obj is None:
            first_op_obj = op_obj

        curr = op_obj
        depth = 0
        while curr and depth < 10:
            depth += 1
            try:
                import NXOpen.CAM
                view_geom = NXOpen.CAM.CAMSetup.View.Geometry
                parent = None
                try:
                    parent = curr.GetParent(view_geom)
                except Exception:
                    try:
                        parent = curr.GetParent()
                    except Exception:
                        parent = None

                if not parent:
                    break
                if is_mcs_geometry_node(parent):
                    mcs_node = parent
                    break
                curr = parent
            except Exception:
                break
        if mcs_node is not None:
            break

    # 備援：若無工步回溯，自 Geometry Root 搜尋第一個 OrientGeometry
    if not mcs_node:
        try:
            import NXOpen.CAM
            geom_root = cam_setup.GetRoot(NXOpen.CAM.CAMSetup.View.Geometry)
            if geom_root:
                def scan_mcs(node):
                    if is_mcs_geometry_node(node):
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

    # 3. 讀取 MCS 資訊 (原點向量、姿態軸向與夾具偏置值)
    g_code = "G54"
    origin = [0.0, 0.0, 0.0]
    x_vec = [1.0, 0.0, 0.0]
    y_vec = [0.0, 1.0, 0.0]
    z_vec = [0.0, 0.0, 1.0]
    csys_tag = None
    has_origin = False

    try:
        # 確保 CAM Session 初始化
        if uf_session and hasattr(uf_session, "Cam"):
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
                
                origin, x_vec, y_vec, z_vec, csys_tag = get_mcs_geometry_axes(
                    builder, mcs_node, uf_session=uf_session
                )
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

    # 4. 依照方案 2 階梯引擎將目標實體投影至加工座標系 (MCS)
    found_box, local_min, local_max, source = extract_target_bounding_box_in_mcs(
        work_part=work_part,
        cam_setup=cam_setup,
        uf_session=uf_session,
        origin=origin,
        x_vec=x_vec,
        y_vec=y_vec,
        z_vec=z_vec,
        csys_tag=csys_tag,
        first_op_obj=first_op_obj,
        mcs_node=mcs_node
    )

    if not found_box or not local_min or not local_max:
        # 若無法讀取邊界盒，預設四面分中頂面碰刀
        return f"加工原點【{g_code}】：X=MID, Y=MID, Z=TOP"

    # 5. 加工座標系局部幾何比對運算
    # 在 MCS 局部空間中，原點自身的坐標永遠是 (0.0, 0.0, 0.0)！
    # local_min[2] 為工件在加工方向的底面位置，local_max[2] 為工件在加工方向的頂面位置
    x_orient = determine_axis_orientation(0.0, local_min[0], local_max[0], is_z_axis=False)
    y_orient = determine_axis_orientation(0.0, local_min[1], local_max[1], is_z_axis=False)
    z_orient = determine_axis_orientation(0.0, local_min[2], local_max[2], is_z_axis=True)

    return f"加工原點【{g_code}】：X={x_orient}, Y={y_orient}, Z={z_orient}"
