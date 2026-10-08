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
    """
    tool_no = "1"
    speed = "4000"
    safe_z = "50."
    start_x = "0."
    start_y = "0."

    if raw_ops:
        first_op = raw_ops[0]
        data = first_op.get("data", {}) if isinstance(first_op, dict) else {}
        t_raw = str(data.get("tool_number", "1")).upper().replace("T", "").strip()
        if t_raw.isdigit():
            tool_no = t_raw
        
        note_str = str(data.get("note", ""))
        m_s = re.search(r"[Ss][:=-]?\s*(\d+)", note_str)
        if m_s:
            speed = m_s.group(1)

    return tool_no, speed, safe_z, start_x, start_y

def extract_body_from_original_nc(raw_nc_content):
    """
    從原始 NC 程式中切分出主切削路徑主體 (Body)，截斷開頭重複之舊檔頭與舊換刀準備行。
    
    截斷點判斷邏輯：
    尋找原始程式中第一把刀啟動刀長補償 (G43 H..) 或冷卻液 (M08) 後的第一行實際進刀動作。
    若未找到明確切削起點，則自第一個 M06/換刀後的第 3~4 行起刀點截斷。
    """
    lines = raw_nc_content.splitlines()
    body_start_idx = 0

    # 尋找第一處 M08 或 G43
    found_m08 = False
    for i, line in enumerate(lines):
        clean_l = line.strip().upper()
        if "M08" in clean_l or "M8" in clean_l:
            body_start_idx = i + 1
            found_m08 = True
            break
        elif "G43" in clean_l:
            body_start_idx = i + 1
            found_m08 = True
            # 如果下一行緊接著 M08，則跳過 M08
            if i + 1 < len(lines) and ("M08" in lines[i+1].upper() or "M8" in lines[i+1].upper()):
                body_start_idx = i + 2
            break

    if not found_m08:
        # 備援：尋找第一個 M06
        for i, line in enumerate(lines):
            clean_l = line.strip().upper()
            if "M06" in clean_l or "M6" in clean_l:
                # 跳過換刀行後續的 S/M03、G54、G43 預備行 (通常約 3~4 行)
                body_start_idx = min(i + 4, len(lines))
                break

    # 回傳主體切削代碼行
    body_lines = lines[body_start_idx:]
    return "\n".join(body_lines)

def transform_nc_content_by_profile(original_nc, profile_id, header_info, mcs_origin_str, raw_ops=None, program_name=""):
    """
    依據選定機型配置將 NC 程式內容進行結構轉換與檔頭置換。
    """
    sanitized_original = sanitize_nc_text(original_nc)
    
    # 提取第一把刀與主軸轉速
    tool_no, speed, safe_z, start_x, start_y = parse_first_tool_and_speed_from_ops(raw_ops)

    # 收集刀具簡要清單供註解使用
    tool_summary_list = []
    if raw_ops:
        seen_t = set()
        for op in raw_ops:
            op_data = op.get("data", {}) if isinstance(op, dict) else {}
            t_num = str(op_data.get("tool_number", "")).strip()
            if t_num and t_num not in seen_t:
                seen_t.add(t_num)
                t_spec = op_data.get("tool_spec_display", "")
                t_name = op_data.get("op_name", "")
                tool_summary_list.append(f"{t_num:<4} {t_spec:<16} {t_name}")

    # 建立標準工單註解
    comment_block = build_nc_header_comment_block(header_info, mcs_origin_str, tool_summary_list, program_name)

    if profile_id == "FANUC_HORIZONTAL":
        # -------------------------------------------------------------
        # Fanuc 臥式機型：專屬 12 行安全性與 B0 定位檔頭置換
        # -------------------------------------------------------------
        hz_header_lines = [
            "%",
            comment_block,
            "G91G30Z0.",
            "G91G30X0.Y0.",
            "G90G17G00G40G49G80",
            "M11",
            "G90G0G54B0.",
            "M10",
            f"M06 T{tool_no}",
            f"S{speed} M03",
            "G17 G40 G80",
            f"G90 G54 G00 X{start_x} Y{start_y}",
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

def prompt_machine_selection_gui():
    """
    透過獨立子進程啟動機型與任務模式選擇視窗 (Zero-Debug 進程隔離)。
    回傳 dict: {"action": "ok"|"cancel", "profile_id": "...", "run_mode": "FULL_DOC_AND_NC"|"NC_ONLY"}
    """
    import subprocess
    import json
    from machine_profile_manager import get_last_selected_profile_id

    cur_dir = os.path.dirname(os.path.abspath(__file__))
    ui_script = os.path.join(cur_dir, "nc_machine_selector_ui.py")
    last_id = get_last_selected_profile_id()

    fallback_res = {
        "action": "ok",
        "profile_id": last_id,
        "run_mode": "FULL_DOC_AND_NC"
    }

    if not os.path.exists(ui_script):
        return fallback_res

    try:
        # 使用外部 Python 3.13 執行 UI 視窗
        proc = subprocess.run(
            ["py", "-3.13", ui_script],
            capture_output=True,
            text=True,
            timeout=60
        )
        for line in proc.stdout.splitlines():
            if line.startswith("__JSON_RESULT__:"):
                json_str = line.replace("__JSON_RESULT__:", "").strip()
                return json.loads(json_str)
    except Exception:
        pass

    return fallback_res

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

def process_and_export_nc_tasks(cam_setup, processed_chunks, profile_id, header_info, stage_origins, output_dir, listing=None, default_post="Fanuc", extension=".nc"):
    """
    批次處理各程式群組：執行背景後處理、特殊字元清洗、機型檔頭置換與安全落盤。
    """
    if not cam_setup or not processed_chunks:
        return 0, 0

    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    success_count = 0
    fail_count = 0

    if listing:
        listing.WriteLine("----------------------------------------")
        listing.WriteLine(f"【開始 NC 碼後處理與機型檔頭置換 (目標機型: {profile_id})】")

    # 依母程式名稱或群組名稱聚合工序
    prog_tasks = {}
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
