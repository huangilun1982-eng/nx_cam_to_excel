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

from machine_profile_manager import (
    find_template_post_files,
    load_available_postprocessors
)


class NcPostDialogApp:
    def __init__(self, master, config):
        self.master = master
        self.config = config

        persisted_settings = {}
        try:
            from machine_profile_manager import get_post_settings
            persisted_settings = get_post_settings()
        except Exception:
            pass

        def_post = config.get("default_post") or persisted_settings.get("last_selected_post", "Fanuc_2026")
        def_profile = config.get("default_machine_profile") or persisted_settings.get("last_selected_id", "FANUC_HORIZONTAL")
        def_custom_path = config.get("custom_post_path") or persisted_settings.get("last_custom_post_path", "")

        self.default_post = str(def_post).strip()
        self.default_machine_profile = str(def_profile).strip()
        self.custom_post_path = str(def_custom_path).strip()

        self.result = {
            "action": "cancel",
            "postprocessor_name": "",
            "machine_profile": self.default_machine_profile,
            "custom_post_path": self.custom_post_path,
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

        # 視窗初始尺寸與置中 (容納機型規格與四動作按鈕，避免右側按鈕被裁切)
        win_w = 700
        win_h = 610
        screen_w = self.master.winfo_screenwidth()
        screen_h = self.master.winfo_screenheight()
        pos_x = max(20, (screen_w - win_w) // 2)
        pos_y = max(20, (screen_h - win_h) // 2)
        self.master.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        self.master.minsize(680, 540)

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

        # 填入任務並依目前副檔名渲染預覽檔名
        self.refresh_task_listbox()

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

        # 檢驗並恢復上次自訂後處理路徑
        if self.custom_post_path and os.path.exists(self.custom_post_path):
            post_base = os.path.splitext(os.path.basename(self.custom_post_path))[0]
            custom_label = f"[自訂] {post_base}"
            if custom_label not in post_names:
                post_names.insert(0, custom_label)

        if self.default_post and self.default_post not in post_names:
            post_names.insert(0, self.default_post)

        self.combo_post = ttk.Combobox(
            row_post,
            textvariable=self.post_var,
            values=post_names,
            font=("Microsoft JhengHei", 9),
            state="normal"
        )
        self.combo_post.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        # 預設選中上次選定後處理器、Fanuc_2026 或第一項
        if self.default_post in post_names:
            self.combo_post.set(self.default_post)
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
            "Fanuc 臥式機 (B0/G30/四軸專用)": "FANUC_HORIZONTAL",
            "Fanuc 立式機 (標準三軸)": "FANUC_VERTICAL",
            "BROTHER 機台 (攻牙機/小型加工中心)": "BROTHER",
            "原始輸出 (不套用檔頭替換)": "RAW"
        }
        # 同時建立大寫與小寫鍵值的反向映射，確保 100% 雙向相容
        self.profile_reverse_map = {}
        for k, v in self.profile_display_map.items():
            self.profile_reverse_map[v] = k
            self.profile_reverse_map[v.lower()] = k

        profile_display_names = list(self.profile_display_map.keys())
        self.combo_profile = ttk.Combobox(
            row_prof,
            textvariable=self.profile_var,
            values=profile_display_names,
            font=("Microsoft JhengHei", 9),
            state="readonly"
        )
        self.combo_profile.pack(side=tk.LEFT, fill=tk.X, expand=True)

        default_prof_id = self.default_machine_profile
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

        init_dir_raw = self.config.get("output_dir") or os.getcwd()
        self.dir_var = tk.StringVar(value=self._normalize_dir_path(init_dir_raw) or os.getcwd())
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

        entry_dir = tk.Entry(
            row_dir,
            textvariable=self.dir_var,
            font=("Microsoft JhengHei", 9),
            bg="white"
        )
        entry_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

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

        init_ext = str(self.config.get("extension") or ".nc").strip() or ".nc"
        if not init_ext.startswith("."):
            init_ext = "." + init_ext
        self.ext_var = tk.StringVar(value=init_ext)
        combo_ext = ttk.Combobox(
            row_ext,
            textvariable=self.ext_var,
            values=[".nc", ".i", ".ptp", ".mpf", ".h", ".tap", ".txt"],
            font=("Microsoft JhengHei", 9),
            width=10
        )
        combo_ext.pack(side=tk.LEFT)
        self.ext_var.trace_add("write", lambda *_: self.refresh_task_listbox())
        self.refresh_task_listbox()

        # 底部動作按鈕列 (before=main_frame 確保高縮放比例下按鈕列不被壓縮裁切)
        btn_bar = tk.Frame(self.master, bg="#EAEEF3", padx=14, pady=10)
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM, before=main_frame)

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

    def _normalize_dir_path(self, raw_path):
        """清洗使用者輸入或系統對話框回傳之路徑字串，統一為作業系統標準分隔符"""
        cleaned = str(raw_path or "").strip().strip('"').strip("'").strip()
        if not cleaned:
            return ""
        return os.path.normpath(cleaned)

    def _resolve_initial_dir(self, raw_path):
        """解析目前存在之起始目錄供 askdirectory 使用，避免無效路徑導致開啟位置錯誤"""
        candidate = self._normalize_dir_path(raw_path)
        while candidate:
            if os.path.isdir(candidate):
                return candidate
            parent = os.path.dirname(candidate)
            if not parent or parent == candidate:
                break
            candidate = parent
        return os.getcwd()

    def _open_native_dialog(self, dialog_fn, **kwargs):
        """暫時解除置頂並指定 parent 開啟系統原生對話框，防止被 -topmost 視窗遮蔽或失去焦點"""
        was_topmost = False
        try:
            was_topmost = str(self.master.attributes("-topmost")).lower() in ("1", "true")
        except Exception:
            was_topmost = False

        kwargs.setdefault("parent", self.master)
        try:
            if was_topmost:
                self.master.attributes("-topmost", False)
            self.master.update_idletasks()
            return dialog_fn(**kwargs)
        finally:
            try:
                if was_topmost:
                    self.master.attributes("-topmost", True)
                self.master.lift()
                self.master.focus_force()
            except Exception:
                pass

    def browse_custom_post(self):
        """瀏覽本機或自訂後處理器檔案 (.tcl 或 .def)"""
        path = self._open_native_dialog(
            filedialog.askopenfilename,
            title="選擇 NX 後處理器檔案",
            filetypes=[("Postprocessor files", "*.tcl;*.def;*.pui"), ("All files", "*.*")]
        )
        if path:
            norm_path = self._normalize_dir_path(path)
            post_base = os.path.splitext(os.path.basename(norm_path))[0]
            custom_label = f"[自訂] {post_base}"
            # 存入自訂路徑
            self.result["custom_post_path"] = norm_path
            vals = list(self.combo_post["values"])
            if custom_label not in vals:
                vals.insert(0, custom_label)
                self.combo_post["values"] = vals
            self.post_var.set(custom_label)

    def _save_current_settings(self, prof_id, post_name):
        """持久化儲存當前設定至 nc_machine_settings.json"""
        try:
            from machine_profile_manager import save_post_settings
            save_post_settings(
                profile_id=prof_id,
                post_name=post_name if post_name else None,
                custom_post_path=self.result.get("custom_post_path", "")
            )
        except Exception:
            pass

    def browse_output_dir(self):
        """選擇 NC 檔案輸出目錄"""
        init_dir = self._resolve_initial_dir(self.dir_var.get())
        dir_selected = self._open_native_dialog(
            filedialog.askdirectory,
            title="選擇 NC 程式碼輸出目錄",
            initialdir=init_dir
        )
        if dir_selected:
            self.dir_var.set(self._normalize_dir_path(dir_selected))

    def _normalize_extension(self):
        """正規化副檔名字串，空值預設為 .nc 並確保前導點號"""
        raw_ext = self.ext_var.get().strip() if hasattr(self, "ext_var") else str(self.config.get("extension") or ".nc").strip()
        if not raw_ext:
            return ".nc"
        return raw_ext if raw_ext.startswith(".") else "." + raw_ext

    def refresh_task_listbox(self):
        """依目前設定的副檔名更新待轉出任務清單預覽"""
        if not hasattr(self, "task_listbox"):
            return
        ext = self._normalize_extension()
        self.task_listbox.delete(0, tk.END)
        for idx, task in enumerate(self.nc_tasks, 1):
            p_name = task.get("program_name", f"PROG_{idx}")
            op_cnt = task.get("op_count", len(task.get("operations", [])))
            ops_preview = ", ".join(task.get("op_names", [])[:3])
            if op_cnt > 3:
                ops_preview += "..."
            item_text = f" [{idx:02d}] {p_name}{ext}   ({op_cnt} 道工序: {ops_preview})"
            self.task_listbox.insert(tk.END, item_text)

    def _sync_custom_post_path(self, selected_post):
        """選用非自訂後處理器時清除殘留的自訂路徑，避免誤套用舊檔"""
        if selected_post and not selected_post.startswith("[自訂]"):
            self.result["custom_post_path"] = ""

    def _submit_post_action(self, action_name):
        """共用的後處理動作驗證、欄位組裝與設定保存流程"""
        selected_post = self.post_var.get().strip()
        if not selected_post:
            messagebox.showwarning("提示", "請選擇或輸入後處理機台名稱！", parent=self.master)
            return

        output_dir = self._normalize_dir_path(self.dir_var.get())
        if not output_dir:
            messagebox.showwarning("提示", "請指定有效的 NC 輸出目錄！", parent=self.master)
            return

        prof_label = self.profile_var.get().strip()
        prof_id = self.profile_display_map.get(prof_label, "FANUC_HORIZONTAL")

        self._sync_custom_post_path(selected_post)
        self.result["action"] = action_name
        self.result["postprocessor_name"] = selected_post
        self.result["machine_profile"] = prof_id
        self.result["output_dir"] = output_dir
        self.result["extension"] = self._normalize_extension()
        self._save_current_settings(prof_id, selected_post)
        self.master.destroy()

    def on_post_and_export(self):
        """使用者確認轉出 NC 碼並匯出工單"""
        self._submit_post_action("post_and_export")

    def on_post_only(self):
        """使用者選擇【⚡ 只轉 NC 碼 (不轉工單)】"""
        self._submit_post_action("post_only")

    def on_export_only(self):
        """使用者選擇僅匯出工單 (不轉 NC)"""
        prof_label = self.profile_var.get().strip()
        prof_id = self.profile_display_map.get(prof_label, "FANUC_HORIZONTAL")
        selected_post = self.post_var.get().strip()

        self._sync_custom_post_path(selected_post)
        self.result["action"] = "export_only"
        self.result["machine_profile"] = prof_id
        if selected_post:
            self.result["postprocessor_name"] = selected_post
        self._save_current_settings(prof_id, selected_post)
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
