# -*- coding: utf-8 -*-
"""
工單分頁預覽與排版決策對話視窗 (獨立進程組件)
職責單純原則：
  1. 接收各工段刀具工步數量，精確預先計算各排版模式之「各工段頁數」與「工單總頁數」。
  2. 呈現精緻現代化預覽視窗，清晰向使用者預告各模式之頁數規劃明細。
  3. 提供直觀明確的決策按鈕，按鈕上直接標明該選項產生之「總頁數」：
     - 【📑 採用單頁延伸 (共 X 頁)】：刀具向下增加格數延伸，示圖同頁縮小或專屬大圖頁。
     - 【📋 採用自動分頁 (共 Y 頁)】：依每頁標準 10 格分頁，每頁皆包含刀具與加工示圖。
     - 【❌ 取消】：中止操作。
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
            "mode": "extend"  # 預設為單頁延伸 (精簡緊湊)
        }
        self.stage_stats = config.get("stage_stats", [])
        self.total_extend = config.get("total_extend_pages", 1)
        self.total_paginate = config.get("total_paginate_pages", 1)
        self.setup_ui()

    def setup_ui(self):
        self.master.title("工單分頁預覽與排版決策")
        self.master.attributes("-topmost", True)
        self.master.resizable(False, False)
        self.master.configure(bg="#F4F6F9")

        win_w = 580
        win_h = 430
        screen_w = self.master.winfo_screenwidth()
        screen_h = self.master.winfo_screenheight()
        pos_x = max(20, (screen_w - win_w) // 2)
        pos_y = max(20, (screen_h - win_h) // 2)
        self.master.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

        # 頂部提示欄
        top_frame = tk.Frame(self.master, bg="#004B87", padx=18, pady=12)
        top_frame.pack(fill=tk.X)

        lbl_header = tk.Label(
            top_frame,
            text="📑 工單分頁預覽與排版決策",
            font=("Microsoft JhengHei", 12, "bold"),
            fg="white",
            bg="#004B87"
        )
        lbl_header.pack(anchor="w")

        lbl_sub = tk.Label(
            top_frame,
            text="系統已預先計算各工段排版頁數，請預覽規劃並決定工單呈現方式：",
            font=("Microsoft JhengHei", 9),
            fg="#DCE6F1",
            bg="#004B87"
        )
        lbl_sub.pack(anchor="w", pady=(2, 0))

        # 內容清單區
        content_frame = tk.Frame(self.master, bg="#F4F6F9", padx=16, pady=10)
        content_frame.pack(fill=tk.BOTH, expand=True)

        lbl_list_title = tk.Label(
            content_frame,
            text="【各工段刀具與分頁規劃預覽】",
            font=("Microsoft JhengHei", 9, "bold"),
            fg="#222222",
            bg="#F4F6F9"
        )
        lbl_list_title.pack(anchor="w", pady=(0, 4))

        list_frame = tk.Frame(content_frame, bg="white", bd=1, relief=tk.SOLID)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 9),
            bg="white",
            fg="#1A1A1A",
            bd=0,
            highlightthickness=0,
            selectbackground="#E8F0FE",
            selectforeground="#004B87"
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        scrollbar.config(command=self.listbox.yview)

        for item in self.stage_stats:
            stg = item.get("stage", "通用工段")
            cnt = item.get("count", 0)
            ext_p = item.get("extend_pages", 1)
            pag_p = item.get("paginate_pages", 1)
            ext_desc = item.get("extend_desc", f"{ext_p}頁")
            pag_desc = item.get("paginate_desc", f"{pag_p}頁")

            line_main = f"🔹 工段【{stg}】：共 {cnt} 道刀具工步"
            line_ext = f"     • 📑 單頁延伸規劃 ➔ {ext_desc}"
            line_pag = f"     • 📋 自動分頁規劃 ➔ {pag_desc}"
            self.listbox.insert(tk.END, line_main)
            self.listbox.insert(tk.END, line_ext)
            self.listbox.insert(tk.END, line_pag)
            self.listbox.insert(tk.END, "")

        # 總頁數對比預覽卡片
        summary_card = tk.Frame(content_frame, bg="#FFFFFF", bd=1, relief=tk.SOLID, padx=12, pady=8)
        summary_card.pack(fill=tk.X, pady=(0, 4))

        lbl_card_title = tk.Label(
            summary_card,
            text="📊 工單總頁數即時對比預告：",
            font=("Microsoft JhengHei", 9, "bold"),
            fg="#333333",
            bg="#FFFFFF"
        )
        lbl_card_title.pack(anchor="w", pady=(0, 2))

        comp_box = tk.Frame(summary_card, bg="#FFFFFF")
        comp_box.pack(fill=tk.X)

        lbl_ext_sum = tk.Label(
            comp_box,
            text=f"📑 單頁延伸模式：預計共 【 {self.total_extend} 頁 】",
            font=("Microsoft JhengHei", 10, "bold"),
            fg="#2E7D32",
            bg="#E8F5E9",
            padx=10,
            pady=4,
            bd=1,
            relief=tk.SOLID
        )
        lbl_ext_sum.pack(side=tk.LEFT, padx=(0, 10))

        lbl_pag_sum = tk.Label(
            comp_box,
            text=f"📋 自動分頁模式：預計共 【 {self.total_paginate} 頁 】",
            font=("Microsoft JhengHei", 10, "bold"),
            fg="#0D47A1",
            bg="#E3F2FD",
            padx=10,
            pady=4,
            bd=1,
            relief=tk.SOLID
        )
        lbl_pag_sum.pack(side=tk.LEFT)

        # 底部按鈕欄
        btn_bar = tk.Frame(self.master, bg="#EAEEF3", padx=16, pady=12)
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # 按鈕 1：單頁延伸 (明確標註總頁數)
        btn_extend_text = f" 📑 採用單頁延伸 (共 {self.total_extend} 頁) "
        btn_extend = tk.Button(
            btn_bar,
            text=btn_extend_text,
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#28A745",
            fg="white",
            activebackground="#218838",
            activeforeground="white",
            padx=12,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_extend
        )
        btn_extend.pack(side=tk.LEFT, padx=(0, 8))

        # 按鈕 2：自動分頁 (明確標註總頁數)
        btn_paginate_text = f" 📋 採用自動分頁 (共 {self.total_paginate} 頁) "
        btn_paginate = tk.Button(
            btn_bar,
            text=btn_paginate_text,
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#0078D7",
            fg="white",
            activebackground="#005A9E",
            activeforeground="white",
            padx=12,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_paginate
        )
        btn_paginate.pack(side=tk.LEFT, padx=(0, 8))

        # 按鈕 3：取消
        btn_cancel = tk.Button(
            btn_bar,
            text=" ❌ 取消 ",
            font=("Microsoft JhengHei", 9),
            bg="#D0D5DD",
            fg="#333333",
            activebackground="#B0B5BD",
            activeforeground="black",
            padx=10,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_cancel
        )
        btn_cancel.pack(side=tk.RIGHT)

        self.master.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def on_extend(self):
        self.result["mode"] = "extend"
        self.master.destroy()

    def on_paginate(self):
        self.result["mode"] = "paginate"
        self.master.destroy()

    def on_cancel(self):
        self.result["mode"] = "cancel"
        self.master.destroy()


def parse_args():
    parser = argparse.ArgumentParser(description="工單分頁預覽與排版決策對話視窗")
    parser.add_argument("--cfg-file", type=str, default="", help="輸入設定檔 (JSON 格式)")
    parser.add_argument("--res-file", type=str, default="", help="輸出結果檔 (JSON 格式)")
    parser.add_argument("--auto-action", type=str, default="", help="單元測試自動動作 (paginate, extend, cancel)")
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
        elif args.auto_action == "cancel":
            root.after(50, app.on_cancel)

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
