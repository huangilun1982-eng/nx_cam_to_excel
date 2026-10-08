# -*- coding: utf-8 -*-
"""
組件名稱：nc_machine_selector_ui.py
功能職責：CNC 機型與輸出模式互動選擇器 (獨立子進程 GUI)
設計原則：
1. 進程隔離防護 (Zero-Debug Protocol)：透過獨立 Python 進程執行，絕不干擾或阻塞 NX 主線程。
2. 現代化簡潔介面：支援機型單選、模式單選 (工單+NC vs 僅轉NC)，自動記憶上次配置。
3. 輸出規範：標準輸出單行 JSON，方便呼叫端獲取選擇結果。
"""

import sys
import json
import tkinter as tk
from tkinter import ttk

from machine_profile_manager import (
    load_machine_profiles,
    get_last_selected_profile_id,
    save_last_selected_profile_id
)

def run_selector_dialog():
    profiles = load_machine_profiles()
    last_id = get_last_selected_profile_id()

    selected_result = {
        "action": "cancel",
        "profile_id": last_id,
        "run_mode": "FULL_DOC_AND_NC" # "FULL_DOC_AND_NC" 或 "NC_ONLY"
    }

    root = tk.Tk()
    root.title("NX CAM - 機型與輸出模式設定")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    # 視窗居中
    win_w, win_h = 420, 310
    scr_w = root.winfo_screenwidth()
    scr_h = root.winfo_screenheight()
    pos_x = (scr_w - win_w) // 2
    pos_y = (scr_h - win_h) // 2
    root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

    # 主容器
    main_frame = ttk.Frame(root, padding=16)
    main_frame.pack(fill="both", expand=True)

    # 標題
    lbl_title = ttk.Label(main_frame, text="【CNC 機型與 NC 檔頭配置】", font=("Microsoft JhengHei", 12, "bold"))
    lbl_title.pack(anchor="w", pady=(0, 10))

    # 1. 機型選擇區塊
    grp_machine = ttk.LabelFrame(main_frame, text=" 1. 請選擇目標機型檔頭 ", padding=10)
    grp_machine.pack(fill="x", pady=(0, 10))

    var_profile = tk.StringVar(value=last_id)
    profile_options = [
        ("FANUC_HORIZONTAL", "Fanuc 臥式機 (B0/G30/四軸專用檔頭)"),
        ("FANUC_VERTICAL", "Fanuc 立式機 (標準安全碼 + 工單註解)"),
        ("BROTHER", "BROTHER 機台 (高速攻牙機專用)"),
        ("RAW", "原始輸出 (保留後處理原始碼)")
    ]

    for p_id, p_name in profile_options:
        rb = ttk.Radiobutton(grp_machine, text=p_name, value=p_id, variable=var_profile)
        rb.pack(anchor="w", pady=2)

    # 2. 執行任務模式區塊
    grp_mode = ttk.LabelFrame(main_frame, text=" 2. 執行任務模式 ", padding=10)
    grp_mode.pack(fill="x", pady=(0, 12))

    var_mode = tk.StringVar(value="FULL_DOC_AND_NC")
    rb_full = ttk.Radiobutton(grp_mode, text="完整模式：產出 Excel 工單 ＋ 自動轉出 NC 碼", value="FULL_DOC_AND_NC", variable=var_mode)
    rb_full.pack(anchor="w", pady=2)

    rb_nc_only = ttk.Radiobutton(grp_mode, text="快速模式：僅轉出 NC 碼 (不開 Excel 工單)", value="NC_ONLY", variable=var_mode)
    rb_nc_only.pack(anchor="w", pady=2)

    # 按鈕區塊
    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack(fill="x", pady=(5, 0))

    def on_ok():
        pid = var_profile.get()
        save_last_selected_profile_id(pid)
        selected_result["action"] = "ok"
        selected_result["profile_id"] = pid
        selected_result["run_mode"] = var_mode.get()
        root.destroy()

    def on_cancel():
        selected_result["action"] = "cancel"
        root.destroy()

    btn_ok = ttk.Button(btn_frame, text="確認執行 (Enter)", command=on_ok)
    btn_ok.pack(side="right", padx=(6, 0))
    btn_cancel = ttk.Button(btn_frame, text="取消", command=on_cancel)
    btn_cancel.pack(side="right")

    root.bind("<Return>", lambda event: on_ok())
    root.bind("<Escape>", lambda event: on_cancel())

    root.mainloop()

    # 輸出單行 JSON
    print("__JSON_RESULT__:" + json.dumps(selected_result, ensure_ascii=False))

if __name__ == "__main__":
    run_selector_dialog()
