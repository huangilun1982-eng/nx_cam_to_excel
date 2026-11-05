# -*- coding: utf-8 -*-
"""
CNC 加工示圖精緻置頂拍照按鈕小工具 (獨立進程組件)
特點：
  1. 採用 pythonw.exe 啟動，零黑窗、零 Console 干擾
  2. 視窗小巧精緻 (約 360x135)，置於螢幕最右上角，完全不擋住 NX 中央工件
  3. 支援【F8 畫面擺正】：
     - 啟動後焦點主動歸還 NX 主視窗，鍵盤按 F8 隨時生效
     - 內建【📐 擺正 (F8)】快捷按鈕，滑鼠點擊即可自動擺正
     - 視窗全域監聽 <F8> 快捷鍵，自動轉發 NX 擺正
  4. 配合主腳本 Windows 訊息泵，NX 主視窗滑鼠中鍵旋轉/滾輪縮放完全自由
"""

import sys
import os
import argparse
import tkinter as tk
import ctypes

user32 = ctypes.windll.user32

import json

def parse_args():
    parser = argparse.ArgumentParser(description="CNC 加工示圖精緻拍照小工具")
    parser.add_argument("--stage", type=str, default="通用工段", help="當前工段名稱")
    parser.add_argument("--target-op", type=str, default="", help="該截圖所屬目標工序名稱")
    parser.add_argument("--tools", type=str, default="", help="該工序所屬刀具說明")
    parser.add_argument("--placement", type=str, default="", help="該截圖在工單中的版面位置")
    parser.add_argument("--custom-seq", type=str, default="", help="工段自訂標記")
    parser.add_argument("--index", type=int, default=1, help="當前工段序號")
    parser.add_argument("--total", type=int, default=1, help="總工段數")
    parser.add_argument("--parent-hwnd", type=int, default=0, help="NX 主視窗句柄")
    parser.add_argument("--req-file", type=str, default="", help="擺正請求旗標檔案路徑")
    parser.add_argument("--layer-file", type=str, default="", help="圖層資訊 JSON 路徑")
    parser.add_argument("--layer-apply-file", type=str, default="", help="圖層套用變更 JSON 路徑")
    parser.add_argument("--layer-flag-file", type=str, default="", help="圖層套用旗標路徑")
    return parser.parse_args()

def main():
    args = parse_args()
    stage_name = args.stage
    target_op = args.target_op.strip() if args.target_op else ""
    if not target_op:
        target_op = stage_name if stage_name != "通用工段" else "加工全貌"
    tools_summary = args.tools.strip()
    placement = args.placement.strip()
    custom_seq = args.custom_seq.strip()

    curr_idx = args.index
    total_stages = args.total
    parent_hwnd = args.parent_hwnd
    req_file = args.req_file
    layer_file = args.layer_file
    layer_apply_file = args.layer_apply_file
    layer_flag_file = args.layer_flag_file

    # 若未指定 parent-hwnd，嘗試抓取當前前景視窗 (即呼叫前的 NX 主視窗)
    if parent_hwnd == 0:
        parent_hwnd = user32.GetForegroundWindow()

    display_name = target_op

    root = tk.Tk()
    root.title(f"📷 拍照引導 ({curr_idx}/{total_stages}) - 工序【{target_op}】")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    root.configure(bg="#F0F4F8")

    # 精緻尺寸，置於螢幕右上角 (高度 170 像素清楚呈現工序詳細資訊與 5 個操作按鈕)
    win_w = 520
    win_h = 170
    screen_w = root.winfo_screenwidth()
    pos_x = max(10, screen_w - win_w - 30)
    pos_y = 60
    root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

    exit_code = [2] # 預設 2 為略過或取消

    # 按鈕引用容器 (支援閉包跨作用域即時反饋，統一專案 ref 風格)
    btn_snap_ref = [None]

    def trigger_f8_snap(event=None):
        """觸發 NX 畫面擺正 (F8 貼齊最接近之標準工藝視向)"""
        # 寫入跨進程旗標，由 NX 主進程底層原生執行視圖擺正 (100% 絕對生效且不干擾視窗)
        if req_file:
            try:
                with open(req_file, "w") as f:
                    f.write("snap")
            except Exception:
                pass
        # 提供即時視覺反饋，讓使用者明確感知按鈕已觸發且小工具不關閉
        try:
            if btn_snap_ref[0] is not None:
                btn_snap_ref[0].config(text=" 📐 視角已擺正 ", bg="#004D40")
                root.after(400, lambda: btn_snap_ref[0].config(text=" 📐 視角擺正 (F8) ", bg="#008080") if btn_snap_ref[0] else None)
        except Exception:
            pass

    # 圖層管理視窗
    layer_win_ref = [None]
    layer_vars = {}
    work_layer_set = set()

    def notify_layer_changes():
        """收集勾選狀態並通知主進程即時刷新 NX 視圖"""
        if not layer_apply_file or not layer_flag_file:
            return
        state_dict = {str(k): v.get() for k, v in layer_vars.items()}
        try:
            with open(layer_apply_file, "w", encoding="utf-8") as f:
                json.dump(state_dict, f, ensure_ascii=False)
            with open(layer_flag_file, "w") as f:
                f.write("layer")
        except Exception:
            pass

    def open_layer_manager():
        """打開浮動圖層管理視窗"""
        if layer_win_ref[0] is not None and layer_win_ref[0].winfo_exists():
            layer_win_ref[0].lift()
            return

        top = tk.Toplevel(root)
        top.title("🗂 圖層可見性管理")
        top.attributes("-topmost", True)
        top.geometry(f"340x440+{max(10, pos_x - 350)}+{pos_y}")
        top.configure(bg="#F9FAFB")
        top.bind("<F8>", trigger_f8_snap)
        layer_win_ref[0] = top

        # 頂部快捷控制列
        f_top = tk.Frame(top, bg="#F9FAFB")
        f_top.pack(fill=tk.X, padx=10, pady=8)

        lbl = tk.Label(f_top, text="勾選即時開關圖層：", font=("Microsoft JhengHei", 9, "bold"), bg="#F9FAFB", fg="#333333")
        lbl.pack(side=tk.LEFT)

        def set_all(val):
            for l_k, v in layer_vars.items():
                v.set(val)
            notify_layer_changes()

        btn_none = tk.Button(f_top, text="全隱", font=("Microsoft JhengHei", 8), bg="#E5E7EB", relief=tk.FLAT, command=lambda: set_all(False), padx=8)
        btn_none.pack(side=tk.RIGHT, padx=2)
        btn_all = tk.Button(f_top, text="全顯", font=("Microsoft JhengHei", 8), bg="#E5E7EB", relief=tk.FLAT, command=lambda: set_all(True), padx=8)
        btn_all.pack(side=tk.RIGHT, padx=2)

        # 滾動區域 (自適應寬度 + 滾動條 + 滾輪支援)
        f_container = tk.Frame(top, bg="#FFFFFF", bd=1, relief=tk.SOLID)
        f_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        canvas = tk.Canvas(f_container, bg="#FFFFFF", highlightthickness=0)
        scrollbar = tk.Scrollbar(f_container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg="#FFFFFF")

        win_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        def on_canvas_configure(e):
            canvas.itemconfig(win_id, width=e.width)
        canvas.bind("<Configure>", on_canvas_configure)

        def on_frame_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        scroll_frame.bind("<Configure>", on_frame_configure)

        def on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass
        top.bind("<MouseWheel>", on_mousewheel)

        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 讀取圖層資訊 (嚴格只列出包含物件之圖層)
        raw_layers = []
        if layer_file and os.path.exists(layer_file):
            try:
                with open(layer_file, "r", encoding="utf-8") as f:
                    raw_layers = json.load(f)
            except Exception:
                pass

        layer_vars.clear()
        work_layer_set.clear()

        # 若無任何包含物件之圖層，顯示清楚提示，不假造空圖層清單
        if not raw_layers:
            lbl_empty = tk.Label(scroll_frame, text="（當前零件中無包含幾何物件之圖層）", font=("Microsoft JhengHei", 9), bg="#FFFFFF", fg="#888888", pady=25)
            lbl_empty.pack(fill=tk.X, expand=True)
        else:
            for item in raw_layers:
                l_num = item["layer"]
                l_name = item.get("name", f"圖層 {l_num}")
                l_vis = item.get("visible", True)
                l_cnt = item.get("count", 0)
                is_wk = item.get("is_work", (l_num == 1))
                if is_wk:
                    work_layer_set.add(l_num)

                var = tk.BooleanVar(value=l_vis)
                layer_vars[l_num] = var

                desc = l_name
                if l_cnt > 0:
                    desc += f"  [{l_cnt}件]"

                cb = tk.Checkbutton(
                    scroll_frame, text=desc, variable=var,
                    font=("Microsoft JhengHei", 9), bg="#FFFFFF", activebackground="#F0F4F8",
                    command=notify_layer_changes, anchor="w", padx=6, pady=3
                )
                cb.pack(fill=tk.X, expand=True)

    def on_confirm():
        exit_code[0] = 0 # 0 代表確認拍照
        root.destroy()

    def on_skip():
        exit_code[0] = 2 # 2 代表略過
        root.destroy()

    def on_back():
        """回到上一步：若為首工序則返回分頁排版編輯器；若為後續工序則返回上一工序重拍"""
        exit_code[0] = 3 # 3 代表回到上一步
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_skip)

    # 視窗全域監聽鍵盤 F8 鍵與 Alt+左鍵 (快捷返回上一步)
    root.bind("<F8>", trigger_f8_snap)
    root.bind("<Alt-Left>", lambda e: on_back())

    # 標題列：指明當前拍照步驟
    lbl_title = tk.Label(
        root, text=f"📷 示圖拍照引導 ({curr_idx}/{total_stages})",
        font=("Microsoft JhengHei", 10, "bold"), bg="#F0F4F8", fg="#004B87"
    )
    lbl_title.pack(pady=(6, 2))

    # 工序用途與所屬刀具詳細資訊卡 (清楚顯示該截圖用於哪一個工序)
    info_card = tk.Frame(root, bg="#E8F1F8", bd=1, relief=tk.SOLID, padx=8, pady=4)
    info_card.pack(fill=tk.X, padx=12, pady=(0, 4))

    # 第一行：目標工序名稱與工單版面位置
    target_line = f"🎯 此截圖用於工序：【{target_op}】"
    if custom_seq:
        target_line += f" (標記: {custom_seq})"
    if placement:
        target_line += f"  |  📄 {placement}"

    lbl_target_info = tk.Label(
        info_card, text=target_line,
        font=("Microsoft JhengHei", 9, "bold"), bg="#E8F1F8", fg="#0F4C81", anchor="w"
    )
    lbl_target_info.pack(fill=tk.X)

    # 第二行：所屬刀具規格
    tools_line = f"🔧 包含刀具：{tools_summary}" if tools_summary else f"🔧 包含刀具：工序【{target_op}】所有加工刀軌"
    lbl_tools_info = tk.Label(
        info_card, text=tools_line,
        font=("Microsoft JhengHei", 8), bg="#E8F1F8", fg="#334155", anchor="w"
    )
    lbl_tools_info.pack(fill=tk.X)

    lbl_hint = tk.Label(
        root, text="中鍵旋轉視圖 | 支援 F8 擺正與圖層開關 | ⬅ 可隨時返回上一步",
        font=("Microsoft JhengHei", 8), bg="#F0F4F8", fg="#64748B"
    )
    lbl_hint.pack(pady=(0, 5))

    # 按鈕列
    btn_frame = tk.Frame(root, bg="#F0F4F8")
    btn_frame.pack()

    # 按鈕 1：畫面擺正 (F8)
    btn_snap = tk.Button(
        btn_frame, text=" 📐 視角擺正 (F8) ", command=trigger_f8_snap,
        bg="#008080", fg="white", activebackground="#006666", activeforeground="white",
        font=("Microsoft JhengHei", 9, "bold"), padx=6, pady=4, relief=tk.FLAT, cursor="hand2"
    )
    btn_snap.pack(side=tk.LEFT, padx=3)
    btn_snap_ref[0] = btn_snap

    # 按鈕 2：圖層控制
    btn_layer = tk.Button(
        btn_frame, text=" 🗂 圖層控制 ", command=open_layer_manager,
        bg="#4A5568", fg="white", activebackground="#2D3748", activeforeground="white",
        font=("Microsoft JhengHei", 9, "bold"), padx=5, pady=4, relief=tk.FLAT, cursor="hand2"
    )
    btn_layer.pack(side=tk.LEFT, padx=3)

    # 按鈕 3：回到上一步 (依使用者指示置於「立即拍照」左側)
    # 第一工序顯示「⬅ 返回排版」，後續工序顯示「⬅ 上一步」
    back_btn_text = " ⬅ 返回排版 " if curr_idx == 1 else " ⬅ 上一步 "
    btn_back = tk.Button(
        btn_frame, text=back_btn_text, command=on_back,
        bg="#475569", fg="white", activebackground="#334155", activeforeground="white",
        font=("Microsoft JhengHei", 9, "bold"), padx=6, pady=4, relief=tk.FLAT, cursor="hand2"
    )
    btn_back.pack(side=tk.LEFT, padx=3)

    # 按鈕 4：立即拍照
    btn_confirm = tk.Button(
        btn_frame, text=" 📸 立即拍照 ", command=on_confirm,
        bg="#0078D7", fg="white", activebackground="#005A9E", activeforeground="white",
        font=("Microsoft JhengHei", 9, "bold"), padx=8, pady=4, relief=tk.FLAT, cursor="hand2"
    )
    btn_confirm.pack(side=tk.LEFT, padx=3)

    # 按鈕 4：略過此段
    btn_skip = tk.Button(
        btn_frame, text=" 略過此段 ", command=on_skip,
        bg="#E1E5EA", fg="#444444", activebackground="#D0D5DD", activeforeground="#000000",
        font=("Microsoft JhengHei", 9), padx=6, pady=4, relief=tk.FLAT, cursor="hand2"
    )
    btn_skip.pack(side=tk.LEFT, padx=3)

    # 視窗置頂但絕不奪取焦點，主動將焦點歸還給 NX 主視窗
    root.lift()
    root.after(50, lambda: user32.SetForegroundWindow(parent_hwnd) if parent_hwnd else None)

    root.mainloop()
    sys.exit(exit_code[0])

if __name__ == "__main__":
    try:
        main()
    except Exception:
        sys.exit(1)
