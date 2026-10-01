# -*- coding: utf-8 -*-
"""
工段刀具分頁決策提示對話視窗 (獨立進程組件)
職責單純原則：
  1. 僅在檢測到某工段刀具工步超過 15 格時啟動。
  2. 呈現精緻現代化對話視窗，向使用者清楚說明該工段工步數量。
  3. 提供兩個明確抉擇按鈕：
     - 【📋 自動分頁顯示】：每頁維持標準 10 格分頁，每頁皆包含刀具與加工示圖。
     - 【📑 單頁延伸顯示】：刀具向下延伸；下方示圖縮小，若放不下則示圖獨立移至下一頁。
  4. 採用獨立進程 (pythonw.exe) 運行，透過 JSON 檔案傳遞參數與回傳結果，徹底隔絕環境干擾。
"""

import os
import sys
import json
import argparse
import tkinter as tk
from tkinter import ttk

class PaginationPromptDialogApp:
    def __init__(self, master, config):
        self.master = master
        self.config = config
        self.result = {
            "mode": "paginate"  # 預設為 paginate (分頁顯示)
        }
        self.stage_stats = config.get("stage_stats", [])
        self.setup_ui()

    def setup_ui(self):
        self.master.title("工段刀具分頁設定")
        self.master.attributes("-topmost", True)
        self.master.resizable(False, False)
        self.master.configure(bg="#F4F6F9")

        win_w = 480
        win_h = 320
        screen_w = self.master.winfo_screenwidth()
        screen_h = self.master.winfo_screenheight()
        pos_x = max(20, (screen_w - win_w) // 2)
        pos_y = max(20, (screen_h - win_h) // 2)
        self.master.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

        # 頂部提示欄
        top_frame = tk.Frame(self.master, bg="#004B87", padx=16, pady=12)
        top_frame.pack(fill=tk.X)

        lbl_header = tk.Label(
            top_frame,
            text="📑 工段刀具超過 15 格提示",
            font=("Microsoft JhengHei", 12, "bold"),
            fg="white",
            bg="#004B87"
        )
        lbl_header.pack(anchor="w")

        lbl_sub = tk.Label(
            top_frame,
            text="偵測到以下工段之刀具工步數量較多，請選擇排版方式：",
            font=("Microsoft JhengHei", 9),
            fg="#DCE6F1",
            bg="#004B87"
        )
        lbl_sub.pack(anchor="w", pady=(2, 0))

        # 內容清單區
        content_frame = tk.Frame(self.master, bg="#F4F6F9", padx=16, pady=12)
        content_frame.pack(fill=tk.BOTH, expand=True)

        list_frame = tk.Frame(content_frame, bg="white", bd=1, relief=tk.SOLID)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 10),
            bg="white",
            fg="#1A1A1A",
            bd=0,
            highlightthickness=0
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        scrollbar.config(command=self.listbox.yview)

        for item in self.stage_stats:
            stg = item.get("stage", "通用工段")
            cnt = item.get("count", 0)
            self.listbox.insert(tk.END, f"  🔹 工段【{stg}】：共 {cnt} 道刀具工步 (已超過 15 格)")

        lbl_hint = tk.Label(
            content_frame,
            text="💡 提示：若選擇「單頁延伸」，刀具過多放不下截圖時，系統將自動將截圖移至下一頁。",
            font=("Microsoft JhengHei", 8),
            fg="#666666",
            bg="#F4F6F9"
        )
        lbl_hint.pack(anchor="w")

        # 底部按鈕欄
        btn_bar = tk.Frame(self.master, bg="#EAEEF3", padx=16, pady=12)
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # 按鈕 1：自動分頁顯示 (依標準10格每頁分頁)
        btn_paginate = tk.Button(
            btn_bar,
            text=" 📋 自動分頁顯示\n(標準10格/頁，每頁含示圖) ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#0078D7",
            fg="white",
            activebackground="#005A9E",
            activeforeground="white",
            padx=10,
            pady=4,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_paginate
        )
        btn_paginate.pack(side=tk.LEFT, padx=(0, 8))

        # 按鈕 2：單頁延伸顯示 (截圖縮小或下一頁)
        btn_extend = tk.Button(
            btn_bar,
            text=" 📑 單頁延伸顯示\n(截圖縮小或移至下一頁) ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#28A745",
            fg="white",
            activebackground="#218838",
            activeforeground="white",
            padx=10,
            pady=4,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_extend
        )
        btn_extend.pack(side=tk.RIGHT)

        self.master.protocol("WM_DELETE_WINDOW", self.on_paginate)

    def on_paginate(self):
        self.result["mode"] = "paginate"
        self.master.destroy()

    def on_extend(self):
        self.result["mode"] = "extend"
        self.master.destroy()


def parse_args():
    parser = argparse.ArgumentParser(description="工段刀具分頁提示對話視窗")
    parser.add_argument("--cfg-file", type=str, default="", help="輸入設定檔 (JSON 格式)")
    parser.add_argument("--res-file", type=str, default="", help="輸出結果檔 (JSON 格式)")
    parser.add_argument("--auto-action", type=str, default="", help="單元測試自動動作 (paginate, extend)")
    return parser.parse_args()


def main():
    args = parse_args()
    config = {}

    if args.cfg_file and os.path.exists(args.cfg_file):
        try:
            with open(args.cfg_file, "r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception:
            config = {}

    root = tk.Tk()
    app = PaginationPromptDialogApp(root, config)

    if args.auto_action:
        if args.auto_action == "paginate":
            root.after(50, app.on_paginate)
        elif args.auto_action == "extend":
            root.after(50, app.on_extend)

    root.mainloop()

    if args.res_file:
        try:
            with open(args.res_file, "w", encoding="utf-8") as f:
                json.dump(app.result, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


if __name__ == "__main__":
    try:
        main()
    except Exception:
        sys.exit(1)
