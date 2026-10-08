# -*- coding: utf-8 -*-
"""
NX CAM 後處理設定與互動確認對話視窗 (獨立進程組件)
職責單純原則：
  1. 負責讀取並解析本機 NX template_post.dat 後處理機台清單。
  2. 呈現現代化對話視窗，展示待轉出 NC 檔案與包含工序數量清單。
  3. 讓使用者選擇後處理機台（或自訂瀏覽 .tcl/.def）、指定輸出路徑與副檔名。
  4. 提供三種明確抉擇按鈕：
     - 【轉出 NC 碼並匯出工單】
     - 【僅匯出工單 (不轉 NC)】
     - 【取消】
  5. 採用獨立進程 (pythonw.exe) 運行，透過 JSON 檔案進行參數傳遞與回傳，徹底杜絕 NX Python 與 Tkinter/Tcl 相衝問題。
"""

import os
import sys
import json
import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ctypes

user32 = ctypes.windll.user32 if hasattr(ctypes, "windll") else None


def find_template_post_files():
    """尋找本機 NX 安裝目錄中的 template_post.dat 檔案清單"""
    found_files = []
    
    # 0. 優先加入本機已知實體路徑
    known_paths = [
        r"C:\Siemens\DesigncenterNX2512\MACH\resource\postprocessor\template_post.dat",
        r"C:\Program Files\Siemens\NX2206\MACH\resource\postprocessor\template_post.dat",
    ]
    for kp in known_paths:
        if os.path.exists(kp) and kp not in found_files:
            found_files.append(kp)

    # 1. 從環境變數取得
    ug_cam_post = os.environ.get("UGII_CAM_POST_DIR", "")
    if ug_cam_post:
        cand = os.path.join(ug_cam_post, "template_post.dat")
        if os.path.exists(cand) and cand not in found_files:
            found_files.append(cand)
            
    ug_base = os.environ.get("UGII_BASE_DIR", "")
    if ug_base:
        cand = os.path.join(ug_base, "MACH", "resource", "postprocessor", "template_post.dat")
        if os.path.exists(cand) and cand not in found_files:
            found_files.append(cand)

    return found_files


def load_available_postprocessors():
    """
    解析 template_post.dat，取得機台清單
    支援常見的兩種格式：
      - 3欄位: 機台名, TCL檔, DEF檔
      - 4欄位: 機台名, 格式, TCL檔, DEF檔
    回傳: list of dict [{"name": "Fanuc_2026", "desc": "Fanuc_2026", "tcl": ..., "def": ...}, ...]
    """
    dat_files = find_template_post_files()
    postprocessors = []
    seen_names = set()

    for dat_file in dat_files:
        post_dir = os.path.dirname(dat_file)
        lines = []
        for enc in ["cp950", "utf-8", "ansi"]:
            try:
                with open(dat_file, "r", encoding=enc) as f:
                    lines = [l.strip() for l in f if l.strip() and not l.strip().startswith("#")]
                if lines:
                    break
            except Exception:
                continue

        for line in lines:
            parts = [p.strip() for p in line.split(",")]
            post_name = ""
            tcl_raw = ""
            def_raw = ""
            
            if len(parts) == 3:
                post_name = parts[0]
                tcl_raw = parts[1]
                def_raw = parts[2]
            elif len(parts) >= 4:
                post_name = parts[0]
                tcl_raw = parts[2]
                def_raw = parts[3]
            else:
                continue

            if not post_name or post_name in seen_names:
                continue
            
            # 濾除佔位符或無效名稱
            if "xxxxx" in post_name.lower() or "xxxxx" in tcl_raw.lower():
                continue

            # 解析變數替換
            tcl_path = tcl_raw.replace("${UGII_CAM_POST_DIR}", post_dir + os.sep)
            def_path = def_raw.replace("${UGII_CAM_POST_DIR}", post_dir + os.sep)
            
            # 若為具體檔案路徑且檔案不存在，則跳過
            if os.path.isabs(tcl_path) and not os.path.exists(tcl_path):
                # 容錯：檢查是否副檔名小寫或缺少
                pass

            seen_names.add(post_name)
            postprocessors.append({
                "name": post_name,
                "desc": post_name,
                "tcl": tcl_path,
                "def": def_path,
                "source_dat": dat_file
            })

    return postprocessors


class NcPostDialogApp:
    def __init__(self, master, config):
        self.master = master
        self.config = config
        self.result = {
            "action": "cancel",
            "postprocessor_name": "",
            "machine_profile": "fanuc_horizontal",
            "custom_post_path": "",
            "output_dir": config.get("output_dir", ""),
            "extension": ".nc"
        }

        self.available_posts = load_available_postprocessors()
        self.nc_tasks = config.get("nc_tasks", [])

        self.setup_ui()

    def setup_ui(self):
        self.master.title("NX CAM 後處理轉出與工單設定")
        self.master.attributes("-topmost", True)
        self.master.resizable(True, True)
        self.master.configure(bg="#F4F6F9")

        # 視窗初始尺寸與置中 (容納機型規格與四動作按鈕)
        win_w = 640
        win_h = 600
        screen_w = self.master.winfo_screenwidth()
        screen_h = self.master.winfo_screenheight()
        pos_x = max(20, (screen_w - win_w) // 2)
        pos_y = max(20, (screen_h - win_h) // 2)
        self.master.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        self.master.minsize(600, 520)

        # 頂部提示標題欄
        top_frame = tk.Frame(self.master, bg="#004B87", padx=16, pady=12)
        top_frame.pack(fill=tk.X)

        lbl_header = tk.Label(
            top_frame,
            text="⚙️ NX CAM 後處理轉出與工單匯出",
            font=("Microsoft JhengHei", 12, "bold"),
            fg="white",
            bg="#004B87"
        )
        lbl_header.pack(anchor="w")

        lbl_sub = tk.Label(
            top_frame,
            text="請選擇是否依輸出的表單轉出 NC 碼，並設定後處理機台與輸出目錄：",
            font=("Microsoft JhengHei", 9),
            fg="#DCE6F1",
            bg="#004B87"
        )
        lbl_sub.pack(anchor="w", pady=(2, 0))

        # 內容主容器
        main_frame = tk.Frame(self.master, bg="#F4F6F9", padx=16, pady=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. NC 任務清單區塊
        lbl_list_title = tk.Label(
            main_frame,
            text=f"📁 預計轉出之 NC 程式清單 (共 {len(self.nc_tasks)} 個程式檔)：",
            font=("Microsoft JhengHei", 10, "bold"),
            fg="#333333",
            bg="#F4F6F9"
        )
        lbl_list_title.pack(anchor="w", pady=(0, 4))

        list_frame = tk.Frame(main_frame, bg="white", bd=1, relief=tk.SOLID)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.task_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 10),
            bg="white",
            fg="#1A1A1A",
            selectbackground="#0078D7",
            selectforeground="white",
            bd=0,
            highlightthickness=0
        )
        self.task_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        scrollbar.config(command=self.task_listbox.yview)

        # 填入任務
        for idx, task in enumerate(self.nc_tasks, 1):
            p_name = task.get("program_name", f"PROG_{idx}")
            op_cnt = task.get("op_count", len(task.get("operations", [])))
            ops_preview = ", ".join(task.get("op_names", [])[:3])
            if op_cnt > 3:
                ops_preview += "..."
            item_text = f" [{idx:02d}] {p_name}.nc   ({op_cnt} 道工序: {ops_preview})"
            self.task_listbox.insert(tk.END, item_text)

        # 2. 機台後處理器設定區塊
        post_frame = tk.LabelFrame(
            main_frame,
            text=" 🛠️ 機台後處理器 (Postprocessor) ",
            font=("Microsoft JhengHei", 9, "bold"),
            fg="#004B87",
            bg="#F4F6F9",
            padx=10,
            pady=8
        )
        post_frame.pack(fill=tk.X, pady=(0, 8))

        # 下拉選單列
        row_post = tk.Frame(post_frame, bg="#F4F6F9")
        row_post.pack(fill=tk.X)

        tk.Label(
            row_post,
            text="選擇機台：",
            font=("Microsoft JhengHei", 9),
            bg="#F4F6F9",
            width=10,
            anchor="w"
        ).pack(side=tk.LEFT)

        self.post_var = tk.StringVar()
        post_names = [p["name"] for p in self.available_posts]
        if not post_names:
            post_names = ["Fanuc_2026", "YCM_noCheck", "840D_3axis", "Mazak_AC"]

        self.combo_post = ttk.Combobox(
            row_post,
            textvariable=self.post_var,
            values=post_names,
            font=("Microsoft JhengHei", 9),
            state="normal"
        )
        self.combo_post.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        # 預設選中 Fanuc_2026 或第一項
        default_post = self.config.get("default_post", "Fanuc_2026")
        if default_post in post_names:
            self.combo_post.set(default_post)
        elif post_names:
            self.combo_post.set(post_names[0])

        btn_browse_post = tk.Button(
            row_post,
            text="瀏覽外部...",
            font=("Microsoft JhengHei", 9),
            bg="#E1E5EA",
            relief=tk.GROOVE,
            command=self.browse_custom_post,
            cursor="hand2"
        )
        btn_browse_post.pack(side=tk.RIGHT)

        # 機型檔頭規格下拉選單列
        row_prof = tk.Frame(post_frame, bg="#F4F6F9")
        row_prof.pack(fill=tk.X, pady=(6, 0))

        tk.Label(
            row_prof,
            text="機型檔頭：",
            font=("Microsoft JhengHei", 9),
            bg="#F4F6F9",
            width=10,
            anchor="w"
        ).pack(side=tk.LEFT)

        self.profile_var = tk.StringVar()
        self.profile_display_map = {
            "Fanuc 臥式機 (B0/G30/四軸專用)": "fanuc_horizontal",
            "Fanuc 立式機 (標準三軸)": "fanuc_vertical",
            "BROTHER 機台 (攻牙機/小型加工中心)": "brother",
            "原始輸出 (不套用檔頭替換)": "raw"
        }
        self.profile_reverse_map = {v: k for k, v in self.profile_display_map.items()}

        profile_display_names = list(self.profile_display_map.keys())
        self.combo_profile = ttk.Combobox(
            row_prof,
            textvariable=self.profile_var,
            values=profile_display_names,
            font=("Microsoft JhengHei", 9),
            state="readonly"
        )
        self.combo_profile.pack(side=tk.LEFT, fill=tk.X, expand=True)

        default_prof_id = self.config.get("default_machine_profile", "fanuc_horizontal")
        default_prof_label = self.profile_reverse_map.get(default_prof_id, profile_display_names[0])
        self.combo_profile.set(default_prof_label)

        # 3. 輸出設定區塊 (目錄與副檔名)
        out_frame = tk.LabelFrame(
            main_frame,
            text=" 📂 輸出設定 ",
            font=("Microsoft JhengHei", 9, "bold"),
            fg="#004B87",
            bg="#F4F6F9",
            padx=10,
            pady=8
        )
        out_frame.pack(fill=tk.X, pady=(0, 8))

        # 輸出目錄列
        row_dir = tk.Frame(out_frame, bg="#F4F6F9")
        row_dir.pack(fill=tk.X, pady=(0, 4))

        tk.Label(
            row_dir,
            text="輸出目錄：",
            font=("Microsoft JhengHei", 9),
            bg="#F4F6F9",
            width=10,
            anchor="w"
        ).pack(side=tk.LEFT)

        self.dir_var = tk.StringVar(value=self.config.get("output_dir", os.getcwd()))
        entry_dir = tk.Entry(
            row_dir,
            textvariable=self.dir_var,
            font=("Microsoft JhengHei", 9),
            bg="white"
        )
        entry_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        btn_browse_dir = tk.Button(
            row_dir,
            text="變更目錄...",
            font=("Microsoft JhengHei", 9),
            bg="#E1E5EA",
            relief=tk.GROOVE,
            command=self.browse_output_dir,
            cursor="hand2"
        )
        btn_browse_dir.pack(side=tk.RIGHT)

        # 副檔名列
        row_ext = tk.Frame(out_frame, bg="#F4F6F9")
        row_ext.pack(fill=tk.X)

        tk.Label(
            row_ext,
            text="NC 副檔名：",
            font=("Microsoft JhengHei", 9),
            bg="#F4F6F9",
            width=10,
            anchor="w"
        ).pack(side=tk.LEFT)

        self.ext_var = tk.StringVar(value=".nc")
        combo_ext = ttk.Combobox(
            row_ext,
            textvariable=self.ext_var,
            values=[".nc", ".i", ".ptp", ".mpf", ".h", ".tap", ".txt"],
            font=("Microsoft JhengHei", 9),
            width=10
        )
        combo_ext.pack(side=tk.LEFT)

        # 底部動作按鈕列
        btn_bar = tk.Frame(self.master, bg="#EAEEF3", padx=14, pady=10)
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # 按鈕 1：轉出 NC 碼並匯出工單
        btn_post_export = tk.Button(
            btn_bar,
            text=" 🚀 轉出 NC 碼並匯出工單 ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#0078D7",
            fg="white",
            activebackground="#005A9E",
            activeforeground="white",
            padx=8,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_post_and_export
        )
        btn_post_export.pack(side=tk.LEFT, padx=(0, 6))

        # 按鈕 2：只轉 NC 碼 (不轉工單)
        btn_post_only = tk.Button(
            btn_bar,
            text=" ⚡ 只轉 NC 碼 (不轉工單) ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#00838F",
            fg="white",
            activebackground="#006064",
            activeforeground="white",
            padx=8,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_post_only
        )
        btn_post_only.pack(side=tk.LEFT, padx=(0, 6))

        # 按鈕 3：僅匯出工單 (不轉 NC)
        btn_export_only = tk.Button(
            btn_bar,
            text=" 📋 僅匯出工單 (不轉 NC) ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#28A745",
            fg="white",
            activebackground="#218838",
            activeforeground="white",
            padx=8,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_export_only
        )
        btn_export_only.pack(side=tk.LEFT, padx=(0, 6))

        # 按鈕 4：取消
        btn_cancel = tk.Button(
            btn_bar,
            text=" ❌ 取消 ",
            font=("Microsoft JhengHei", 9),
            bg="#D0D5DD",
            fg="#333333",
            activebackground="#B0B5BD",
            activeforeground="black",
            padx=8,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_cancel
        )
        btn_cancel.pack(side=tk.RIGHT)

        self.master.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def browse_custom_post(self):
        """瀏覽本機或自訂後處理器檔案 (.tcl 或 .def)"""
        path = filedialog.askopenfilename(
            title="選擇 NX 後處理器檔案",
            filetypes=[("Postprocessor files", "*.tcl;*.def;*.pui"), ("All files", "*.*")]
        )
        if path:
            post_base = os.path.splitext(os.path.basename(path))[0]
            # 存入自訂路徑
            self.result["custom_post_path"] = path
            self.post_var.set(f"[自訂] {post_base}")

    def browse_output_dir(self):
        """選擇 NC 檔案輸出目錄"""
        dir_selected = filedialog.askdirectory(
            title="選擇 NC 程式碼輸出目錄",
            initialdir=self.dir_var.get()
        )
        if dir_selected:
            self.dir_var.set(dir_selected)

    def on_post_and_export(self):
        """使用者確認轉出 NC 碼並匯出工單"""
        selected_post = self.post_var.get().strip()
        if not selected_post:
            messagebox.showwarning("提示", "請選擇或輸入後處理機台名稱！", parent=self.master)
            return

        prof_label = self.profile_var.get().strip()
        prof_id = self.profile_display_map.get(prof_label, "fanuc_horizontal")

        self.result["action"] = "post_and_export"
        self.result["postprocessor_name"] = selected_post
        self.result["machine_profile"] = prof_id
        self.result["output_dir"] = self.dir_var.get().strip()
        self.result["extension"] = self.ext_var.get().strip()
        if not self.result["extension"].startswith("."):
            self.result["extension"] = "." + self.result["extension"]
        self.master.destroy()

    def on_post_only(self):
        """使用者選擇【⚡ 只轉 NC 碼 (不轉工單)】"""
        selected_post = self.post_var.get().strip()
        if not selected_post:
            messagebox.showwarning("提示", "請選擇或輸入後處理機台名稱！", parent=self.master)
            return

        prof_label = self.profile_var.get().strip()
        prof_id = self.profile_display_map.get(prof_label, "fanuc_horizontal")

        self.result["action"] = "post_only"
        self.result["postprocessor_name"] = selected_post
        self.result["machine_profile"] = prof_id
        self.result["output_dir"] = self.dir_var.get().strip()
        self.result["extension"] = self.ext_var.get().strip()
        if not self.result["extension"].startswith("."):
            self.result["extension"] = "." + self.result["extension"]
        self.master.destroy()

    def on_export_only(self):
        """使用者選擇僅匯出工單 (不轉 NC)"""
        prof_label = self.profile_var.get().strip()
        prof_id = self.profile_display_map.get(prof_label, "fanuc_horizontal")
        self.result["action"] = "export_only"
        self.result["machine_profile"] = prof_id
        self.master.destroy()

    def on_cancel(self):
        """使用者取消操作"""
        self.result["action"] = "cancel"
        self.master.destroy()


def parse_args():
    parser = argparse.ArgumentParser(description="NX CAM 後處理確認對話視窗")
    parser.add_argument("--cfg-file", type=str, default="", help="輸入設定檔 (JSON 格式)")
    parser.add_argument("--res-file", type=str, default="", help="輸出結果檔 (JSON 格式)")
    parser.add_argument("--auto-action", type=str, default="", help="自動觸發動作 (供單元測試: post_and_export, post_only, export_only, cancel)")
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
    app = NcPostDialogApp(root, config)

    if args.auto_action:
        # 單元測試自動模式
        if args.auto_action == "post_and_export":
            root.after(50, app.on_post_and_export)
        elif args.auto_action == "post_only":
            root.after(50, app.on_post_only)
        elif args.auto_action == "export_only":
            root.after(50, app.on_export_only)
        elif args.auto_action == "cancel":
            root.after(50, app.on_cancel)

    root.mainloop()

    # 寫入回傳結果
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
