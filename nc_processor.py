# -*- coding: utf-8 -*-
"""
組件名稱：nc_processor.py
功能職責：NX CAM 後處理自動執行、特殊字符清洗與多機型檔頭智慧置換引擎
設計原則：
1. 組件功能單純原則：專注於 NC 檔案後處理產出、文字清洗、檔頭辨識與替換，不涉及 Excel 表格繪圖。
2. 零破壞與多階防禦 (Zero-Debug Protocol)：
   - 輸出前自動建立 .bak 備份。
   - 採用暫存檔原子寫入 (Atomic Write) 與強制落盤 (fsync)。
   - 任何例外自動安全回退，絕不產生破損或半截 NC 檔案。
3. 支援模式：
   - FANUC_HORIZONTAL: 臥式機專屬 12 行安全代碼置換 (B0/M11/M10/G30)
   - FANUC_VERTICAL / BROTHER: 保留原後處理 G 碼 + 注入標準工單註解
   - RAW: 原始碼不變
"""

import os
import re
import datetime

# 全形符號轉半形映射表
FULLWIDTH_TO_HALFWIDTH = {
    ord('（'): '(',
    ord('）'): ')',
    ord('【'): '[',
    ord('】'): ']',
    ord('：'): ':',
    ord('，'): ',',
    ord('。'): '.',
    ord('；'): ';',
    ord('　'): ' ',
    ord('—'): '-',
    ord('～'): '~'
}

def sanitize_nc_text(text):
    """
    清洗 NC 程式文字內容，將全形符號轉換為半形，避免 CNC 控制器傳輸報警或語法錯誤。
    """
    if not text:
        return ""
    # 1. 替換全形括號與標點
    sanitized = text.translate(FULLWIDTH_TO_HALFWIDTH)
    return sanitized

def build_nc_header_comment_block(header_info, mcs_origin_str, tools_summary=None, program_name=""):
    """
    產生統一規範的 NC 程式註解表頭區塊。
    將關鍵字與原點格式標準化為機台通用的英數字格式，避免控制器螢幕顯示亂碼或傳輸警報。
    """
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    dwg_name = sanitize_nc_text(header_info.get("drawing_name", ""))
    dwg_no = sanitize_nc_text(header_info.get("drawing_number", ""))
    blank_sz = sanitize_nc_text(header_info.get("blank_size", ""))
    prog_name = sanitize_nc_text(program_name)

    # 將中文加工原點文字轉為機台通用格式：G54 [X=MID, Y=MID, Z=TOP]
    clean_origin = sanitize_nc_text(mcs_origin_str)
    clean_origin = clean_origin.replace("加工原點", "").replace("【", "[").replace("】", "]").replace("：", ": ").strip()
    if not clean_origin.startswith("[") and "G" in clean_origin:
        clean_origin = re.sub(r":?\s*", ": ", clean_origin, count=1)

    lines = [
        "(========================================)",
        f"( PROGRAM NAME : {prog_name:<24} )" if prog_name else "( PROGRAM NAME : MAIN                     )",
        f"( PART NAME    : {dwg_name:<24} )",
        f"( DRAWING NO   : {dwg_no:<24} )" if dwg_no else "( DRAWING NO   : N/A                      )",
        f"( BLANK SIZE   : {blank_sz:<24} )" if blank_sz else "( BLANK SIZE   : N/A                      )",
        f"( MCS ORIGIN   : {clean_origin:<24} )",
        f"( DATE / TIME  : {now_str:<24} )",
        "(----------------------------------------)"
    ]

    if tools_summary:
        lines.append("( TOOL LIST :                            )")
        for t_info in tools_summary:
            t_line = f"(  {t_info:<37} )"
            lines.append(sanitize_nc_text(t_line))
        lines.append("(----------------------------------------)")

    lines.append("(========================================)")
    return "\n".join(lines)

def parse_first_tool_and_speed_from_ops(raw_ops):
    """
    自工序列表提取第一道工序的 T 號、轉速 S 與安全 Z 值。
    相容原生字典結構與巢狀 data 字典，並支援連續刀號格式 (如 T01~T02)。
    """
    tool_no = "1"
    speed = "4000"
    safe_z = "50."
    start_x = "0."
    start_y = "0."

    if raw_ops:
        first_op = raw_ops[0]
        data = first_op.get("data") if (isinstance(first_op, dict) and "data" in first_op) else (first_op if isinstance(first_op, dict) else {})
        t_raw = str(data.get("tool_number", "1")).upper().replace("T", "").strip()
        if "~" in t_raw:
            t_raw = t_raw.split("~")[0].strip()
        if t_raw.isdigit():
            tool_no = str(int(t_raw))
        
        note_str = str(data.get("note", ""))
        m_s = re.search(r"[Ss][:=-]?\s*(\d+)", note_str)
        if m_s:
            speed = m_s.group(1)

    return tool_no, speed, safe_z, start_x, start_y

def parse_nc_initial_params(nc_text, default_wcs="G54", default_tool="1", default_speed="4000", default_safe_z="50."):
    """
    自原始 NC 程式中解析首把刀具之工作座標系 (WCS, 如 G54/G55/G56)、
    起刀切削點 (X, Y 絕對座標)、安全 Z 高度、刀號與主軸轉速。
    保證臥式機型置換檔頭時，絕對不覆蓋掉 CAM 原生算出的真實首刀起刀點座標！
    """
    lines = nc_text.splitlines()
    code_lines = []
    for line in lines:
        clean = line.strip().upper()
        code_only = re.sub(r'\(.*?\)', '', clean).strip()
        code_lines.append(code_only)
        if re.search(r'\bM0?8\b', code_only):
            break
        if len(code_lines) > 60:
            break

    header_code = "\n".join(code_lines)

    # 尋找 M06 換刀行之位置
    m06_idx = 0
    for idx, l in enumerate(code_lines):
        if re.search(r'\bM0?6\b', l):
            m06_idx = idx
            break

    # 1. 解析 WCS (例如 G54, G55, G56, G57, G58, G59, G54.1 P1 等)
    # 優先自 M06 之後尋找，若無則自全表頭尋找
    wcs = default_wcs
    after_m06_text = "\n".join(code_lines[m06_idx:])
    m_wcs = re.search(r'\b(G5[4-9]|G54\.1\s*P\d+)\b', after_m06_text) or re.search(r'\b(G5[4-9]|G54\.1\s*P\d+)\b', header_code)
    if m_wcs:
        wcs = m_wcs.group(1).replace(" ", "")

    # 2. 解析 Tool 號
    tool = default_tool
    m_t = re.search(r'\bT0*(\d+)\b', header_code)
    if m_t:
        tool = m_t.group(1)

    # 3. 解析主軸轉速 S
    speed = default_speed
    m_s = re.search(r'\bS(\d+)\b', header_code)
    if m_s:
        speed = m_s.group(1)

    # 4. 解析安全 Z (優先自 G43 刀長補償行或換刀後 Z 移動抓取)
    safe_z = default_safe_z
    m_z = (
        re.search(r'G43[^\n]*\bZ([-+]?\d*\.?\d+)\b', after_m06_text) or
        re.search(r'\bZ([-+]?\d*\.?\d+)\b[^\n]*G43', after_m06_text) or
        re.search(r'\bZ([-+]?\d*\.?\d+)\b', after_m06_text)
    )
    if m_z:
        safe_z = m_z.group(1)
        if '.' not in safe_z:
            safe_z += '.'

    # 5. 解析起刀點 X, Y 座標 (在 M06 之後排除 G28/G30/G53 歸零干擾)
    start_x = "0."
    start_y = "0."
    non_homing_lines = [l for l in code_lines[m06_idx:] if not any(k in l for k in ('G28', 'G30', 'G53'))]
    non_homing_text = "\n".join(non_homing_lines)

    m_x = re.search(r'\bX([-+]?\d*\.?\d+)\b', non_homing_text)
    if m_x:
        start_x = m_x.group(1)
        if '.' not in start_x:
            start_x += '.'

    m_y = re.search(r'\bY([-+]?\d*\.?\d+)\b', non_homing_text)
    if m_y:
        start_y = m_y.group(1)
        if '.' not in start_y:
            start_y += '.'

    return wcs, tool, speed, safe_z, start_x, start_y

def extract_body_from_original_nc(raw_nc_content):
    """
    從原始 NC 程式中切分出主切削路徑主體 (Body)，截斷開頭重複之舊檔頭與舊換刀準備行。
    
    截斷點判斷邏輯：
    1. 尋找原始程式中首道工序啟動冷卻液 (M08/M8) 的行，自下一行起視為主切削本體。
    2. 若未設定 M08，則自第一把刀啟動刀長補償 (G43) 的下一行起截斷。
    3. 若無上述指令，則自首個 M06 換刀後的第 4 行起截斷。
    4. 濾除括號註解，避免工序說明或刀具備註中包含 M8 等字眼造成誤判。
    """
    lines = raw_nc_content.splitlines()
    body_start_idx = 0
    found_cut = False

    # 1. 優先尋找首道換刀後之第一處 M08 / M8 (冷卻液啟動行)
    for i, line in enumerate(lines):
        clean_l = re.sub(r'\(.*?\)', '', line).strip().upper()
        if re.search(r'\bM0?8\b', clean_l):
            body_start_idx = i + 1
            found_cut = True
            break

    # 2. 備援 A：若無 M08，則尋找第一處 G43 (刀長補償建立行)
    if not found_cut:
        for i, line in enumerate(lines):
            clean_l = re.sub(r'\(.*?\)', '', line).strip().upper()
            if "G43" in clean_l:
                body_start_idx = i + 1
                found_cut = True
                break

    # 3. 備援 B：若無 M08 亦無 G43，則尋找首個 M06 (換刀行) 並跳過預備行
    if not found_cut:
        for i, line in enumerate(lines):
            clean_l = re.sub(r'\(.*?\)', '', line).strip().upper()
            if re.search(r'\bM0?6\b', clean_l):
                body_start_idx = min(i + 4, len(lines))
                found_cut = True
                break

    body_lines = lines[body_start_idx:]
    return "\n".join(body_lines)

def transform_nc_content_by_profile(original_nc, profile_id, header_info, mcs_origin_str, raw_ops=None, program_name=""):
    """
    依據選定機型配置將 NC 程式內容進行結構轉換與檔頭置換。
    """
    sanitized_original = sanitize_nc_text(original_nc)
    
    # 提取第一把刀與主軸轉速 (作為預設後備值)
    op_tool, op_speed, op_safe_z, op_x, op_y = parse_first_tool_and_speed_from_ops(raw_ops)

    # 收集刀具簡要清單供註解使用
    tool_summary_list = []
    if raw_ops:
        seen_t = set()
        for op in raw_ops:
            op_data = op.get("data") if (isinstance(op, dict) and "data" in op) else (op if isinstance(op, dict) else {})
            t_num = str(op_data.get("tool_number", "")).strip()
            if t_num and t_num not in seen_t:
                seen_t.add(t_num)
                t_spec = op_data.get("tool_spec_display", "")
                t_name = op_data.get("op_name", op_data.get("clean_op_name", ""))
                tool_summary_list.append(f"{t_num:<4} {t_spec:<16} {t_name}")

    # 建立標準工單註解
    comment_block = build_nc_header_comment_block(header_info, mcs_origin_str, tool_summary_list, program_name)

    if profile_id == "FANUC_HORIZONTAL":
        # -------------------------------------------------------------
        # Fanuc 臥式機型：專屬 12 行安全性與 B0 定位檔頭置換
        # -------------------------------------------------------------
        # 從 mcs_origin_str 提取工段原點 WCS (作為預設後備值)
        default_wcs = "G54"
        if mcs_origin_str:
            m_stg_wcs = re.search(r'\b(G5[4-9]|G54\.1\s*P\d+)\b', mcs_origin_str)
            if m_stg_wcs:
                default_wcs = m_stg_wcs.group(1).replace(" ", "")

        # 自原始 NC 程式中動態精確解析真實 WCS、起刀點 X/Y、安全 Z、刀號 T 與轉速 S
        nc_wcs, nc_tool, nc_speed, nc_safe_z, nc_start_x, nc_start_y = parse_nc_initial_params(
            sanitized_original,
            default_wcs=default_wcs,
            default_tool=op_tool,
            default_speed=op_speed,
            default_safe_z=op_safe_z
        )

        # 決定最終採用參數 (起刀座標與 WCS 嚴格取自原始 NC 避免撞機，刀具轉速兼顧 raw_ops 顯式設定)
        tool_no = op_tool if (raw_ops and op_tool != "1") else nc_tool
        speed = op_speed if (raw_ops and op_speed != "4000") else nc_speed
        safe_z = nc_safe_z if nc_safe_z != "50." else op_safe_z
        start_x = nc_start_x
        start_y = nc_start_y
        wcs = nc_wcs

        hz_header_lines = [
            "%",
            comment_block,
            "G91G30Z0.",
            "G91G30X0.Y0.",
            "G90G17G00G40G49G80",
            "M11",
            f"G90G0{wcs}B0.",
            "M10",
            f"M06 T{tool_no}",
            f"S{speed} M03",
            "G17 G40 G80",
            f"G90 {wcs} G00 X{start_x} Y{start_y}",
            f"G0 G43 H{tool_no} Z{safe_z}",
            "M08"
        ]
        hz_header_str = "\n".join(hz_header_lines)

        # 截斷原 NC 重複之換刀前導行，保留真正切削路徑主體
        nc_body = extract_body_from_original_nc(sanitized_original)
        final_nc = hz_header_str + "\n" + nc_body.lstrip()
        if not final_nc.rstrip().endswith("%"):
            final_nc = final_nc.rstrip() + "\n%"
        return final_nc

    elif profile_id in ["FANUC_VERTICAL", "BROTHER"]:
        # -------------------------------------------------------------
        # Fanuc 立式 / BROTHER 機型：保留原後處理碼 + 頂部注入工單註解
        # -------------------------------------------------------------
        lines = sanitized_original.splitlines()
        insert_idx = 0
        # 尋找是否開頭有 % 或 O 號
        for i, line in enumerate(lines[:10]):
            clean_l = line.strip().upper()
            if clean_l == "%" or clean_l.startswith("O"):
                insert_idx = i + 1

        new_lines = lines[:insert_idx] + [comment_block] + lines[insert_idx:]
        final_nc = "\n".join(new_lines)
        return final_nc

    else:
        # RAW 模式
        return sanitized_original



def run_nx_postprocess(cam_setup, cam_objects, post_name, output_file):
    """
    呼叫 NX CAM 官方 Postprocess API 將工序背景後處理輸出至檔案。
    """
    if not cam_setup or not cam_objects:
        return False

    try:
        import NXOpen.CAM
        metric_unit = NXOpen.CAM.CAMSetup.OutputUnits.Metric
        out_dir = os.path.dirname(os.path.abspath(output_file))
        if not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        cam_setup.Postprocess(list(cam_objects), post_name, output_file, metric_unit)
        return os.path.exists(output_file) and os.path.getsize(output_file) > 0
    except Exception:
        return False

def process_and_export_nc_tasks(cam_setup, processed_chunks=None, profile_id="FANUC_HORIZONTAL", header_info=None, stage_origins=None, output_dir="", listing=None, default_post="Fanuc", extension=".nc", nc_tasks=None):
    """
    批次處理各程式群組：執行背景後處理、特殊字元清洗、機型檔頭置換與安全落盤。
    支援直接傳入 nc_tasks (保證未合併之 100% 原始工序物件) 或 processed_chunks (自動展開同刀號合併工序)。
    """
    if not cam_setup:
        return 0, 0
    if not processed_chunks and not nc_tasks:
        return 0, 0

    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    header_info = header_info or {}
    stage_origins = stage_origins or {}
    success_count = 0
    fail_count = 0

    if listing:
        listing.WriteLine("----------------------------------------")
        listing.WriteLine(f"【開始 NC 碼後處理與機型檔頭置換 (目標機型: {profile_id})】")

    # 建立輔助查找映射表 (從 processed_chunks 蒐集已解析的工步資訊與刀具資料)
    chunk_ops_by_prog = {}
    chunk_stage_by_prog = {}
    if processed_chunks:
        for chunk in processed_chunks:
            p_name = chunk.get("parent_program", "") or chunk.get("group_name", "MAIN")
            stage = chunk.get("stage", "通用工段")
            if p_name not in chunk_stage_by_prog:
                chunk_stage_by_prog[p_name] = stage
            if p_name not in chunk_ops_by_prog:
                chunk_ops_by_prog[p_name] = []
            for mop in chunk.get("operations", []):
                chunk_ops_by_prog[p_name].append(mop)

    # 聚合各程式群組的工序任務，確保 100% 納入所有工序物件
    prog_tasks = {}

    # 1. 優先使用 nc_tasks (由 CAM 導覽器選取樹構建之任務，工序未合併，保有完整物件列表)
    if nc_tasks:
        for task in nc_tasks:
            p_name = task.get("program_name", "MAIN")
            stage = task.get("stage") or chunk_stage_by_prog.get(p_name, "通用工段")
            raw_ops_in_task = task.get("operations", [])
            ops_data = task.get("ops_data") or chunk_ops_by_prog.get(p_name, [])

            # 若無現成 ops_data 則自 task["operations"] 構建基礎資訊
            if not ops_data and raw_ops_in_task:
                ops_data = []
                for op in raw_ops_in_task:
                    op_n = getattr(op, "Name", "") if hasattr(op, "Name") else str(op)
                    ops_data.append({"op_name": op_n, "_nx_op_obj": op})

            prog_tasks[p_name] = {
                "stage": stage,
                "ops_data": ops_data,
                "nx_ops": list(raw_ops_in_task)
            }

    # 2. 備援模式：自 processed_chunks 重建任務，並透過 _merged_nx_ops 保證所有合併工序皆被取出
    elif processed_chunks:
        for chunk in processed_chunks:
            p_name = chunk.get("parent_program", "") or chunk.get("group_name", "MAIN")
            stage = chunk.get("stage", "通用工段")
            if p_name not in prog_tasks:
                prog_tasks[p_name] = {
                    "stage": stage,
                    "ops_data": [],
                    "nx_ops": []
                }
            for mop in chunk.get("operations", []):
                prog_tasks[p_name]["ops_data"].append(mop)
                # 萃取所有底層工序物件 (包含因同刀號合併被收納於 _merged_nx_ops 之工序)
                sub_ops = mop.get("_merged_nx_ops")
                if sub_ops and isinstance(sub_ops, (list, tuple)):
                    for nx_op in sub_ops:
                        if nx_op and nx_op not in prog_tasks[p_name]["nx_ops"]:
                            prog_tasks[p_name]["nx_ops"].append(nx_op)
                else:
                    nx_op = mop.get("_nx_op_obj")
                    if nx_op and nx_op not in prog_tasks[p_name]["nx_ops"]:
                        prog_tasks[p_name]["nx_ops"].append(nx_op)

    for prog_name, task_info in prog_tasks.items():
        nx_ops = task_info["nx_ops"]
        ops_data = task_info["ops_data"]
        stage = task_info["stage"]
        mcs_origin_str = stage_origins.get(stage, "加工原點【G54】：X=MID, Y=MID, Z=TOP")

        if not nx_ops:
            if listing:
                listing.WriteLine(f"  [略過] 程式群組 [{prog_name}] 無有效工序物件")
            continue

        final_nc_path = os.path.join(output_dir, f"{prog_name}{extension}")
        tmp_raw_nc = os.path.join(output_dir, f"_{prog_name}_raw_tmp.nc")

        try:
            if listing:
                listing.WriteLine(f"  正在處理 [{prog_name}] (共 {len(nx_ops)} 道工序)...")

            # 1. 執行背景後處理輸出暫存原始碼
            ok = run_nx_postprocess(cam_setup, nx_ops, default_post, tmp_raw_nc)
            if not ok or not os.path.exists(tmp_raw_nc):
                fail_count += 1
                if listing:
                    listing.WriteLine(f"    ✖ 後處理器呼叫失敗 [{prog_name}]")
                continue

            # 2. 讀取原始 NC 碼
            raw_content = ""
            with open(tmp_raw_nc, "r", errors="ignore") as f:
                raw_content = f.read()

            # 3. 執行機型檔頭轉換與特殊字元清洗
            final_content = transform_nc_content_by_profile(
                original_nc=raw_content,
                profile_id=profile_id,
                header_info=header_info,
                mcs_origin_str=mcs_origin_str,
                raw_ops=ops_data,
                program_name=prog_name
            )

            # 4. Zero-Debug 原子寫入目標檔案並備份
            safe_write_nc_file(final_nc_path, final_content)
            success_count += 1

            if listing:
                sz = os.path.getsize(final_nc_path) if os.path.exists(final_nc_path) else 0
                listing.WriteLine(f"    ✔ 成功輸出 NC：{final_nc_path} ({sz} 位元組)")

        except Exception as ex:
            fail_count += 1
            if listing:
                listing.WriteLine(f"    ✖ 轉換異常 [{prog_name}]：{str(ex)}")
        finally:
            if os.path.exists(tmp_raw_nc):
                try:
                    os.remove(tmp_raw_nc)
                except Exception:
                    pass

    if listing:
        listing.WriteLine(f"NC 程式處理完畢：成功 {success_count} 個，失敗 {fail_count} 個。")
        listing.WriteLine("----------------------------------------")

    return success_count, fail_count

def safe_write_nc_file(file_path, content, encoding="utf-8"):
    """
    Zero-Debug Protocol 原子寫入 NC 檔案並自動保留 .bak 備份。
    """
    if os.path.exists(file_path):
        bak_path = file_path + ".bak"
        try:
            if os.path.exists(bak_path):
                os.remove(bak_path)
            os.rename(file_path, bak_path)
        except Exception:
            pass

    tmp_path = file_path + ".tmp"
    with open(tmp_path, "w", encoding=encoding, errors="replace") as f:
        f.write(content)
        f.flush()
        os.fsync(f.fileno())

    if os.path.exists(file_path):
        try:
            os.remove(file_path)
        except Exception:
            pass
    os.rename(tmp_path, file_path)
    return True
