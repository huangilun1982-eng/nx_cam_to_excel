# -*- coding: utf-8 -*-
"""
組件名稱：interactive_pagination_dialog.py
功能職責：工單可視化互動式分頁排版編輯器 (獨立進程 GUI 組件)
核心製造邏輯原則：
1. 檔名群組不可分割原則 (Group Integrity)：
   - CNC CAM 加工中，一個 NC 程式檔名 (NC Group) 包含數個工步刀具。
   - 排序與跨頁移動嚴格以「程式檔名為群組」整體移動，絕不拆散單一刀具。
2. 程式檔名清楚辨識：
   - 表格中清楚顯示程式檔名，群組首列顯示「📁 檔名」，同群組後續刀具顯示「  ↳ (同上: 檔名)」，杜絕空白列。
3. 序號欄直接內聯輸入工段標記 (Direct In-cell Editing on Sequence Column)：
   - 使用者可直接雙擊 (或按 F2) 在「序號 / 標記」儲存格中輸入自訂文字 (如 OP10、正面粗銑)。
   - 若使用者有輸入，序號完全依照輸入文字顯示，不須再加流水序號；若清空則恢復風格流水號。
4. 群組截圖決定 (目標為程式群組，非每把刀具，預設不截圖)：
   - 截圖欄位僅在程式群組首行顯示勾選狀態 (☑ 截圖 / ☐ 略過)，同群組後續刀具保持空白。
   - 預設為不勾選 (need_capture = False)。
   - 支援直接單擊截圖欄位快速切換勾選狀態。
5. 兩種示圖排版模式自由切換：
   - A. 【工序後加圖片】(after_ops)：同頁下方預留示圖展示區；未勾選截圖頁面為純刀具清單。
   - B. 【圖片統一置於最後一頁】(last_page)：前面頁面均為純刀具頁，最後統一建立大圖附圖頁。
6. 群組聯動選取：
   - 點擊任一刀具自動聯動全選所屬整個程式群組，視覺回饋明確直觀。
7. 零除錯防護 (Zero-Debug Protocol)：
   - 原子寫入 (Atomic Write + fsync + os.replace)，異常安全攔截。
"""

import os
import sys
import json
import argparse
import copy
import re
from collections import Counter
import tkinter as tk
from tkinter import ttk, messagebox

# 序號風格常數 (依指示精簡：移除工段前綴、兩位數補零、工藝步長)
SEQ_STYLE_PER_PAGE = "per_page"
SEQ_STYLE_GLOBAL = "global_continuous"

# 舊序號風格常數 (保留常數定義確保向後相容與測試完整性)
SEQ_STYLE_STAGE_PREFIX = "stage_prefix"
SEQ_STYLE_ZERO_PAD = "zero_padded"
SEQ_STYLE_STEP_10 = "step_10"

SEQ_STYLE_OPTIONS = [
    (SEQ_STYLE_PER_PAGE, "1. 每頁獨立 (1, 2, 3...) 【標準預設】"),
    (SEQ_STYLE_GLOBAL, "2. 全工單連續 (1, 2... 15, 16)")
]

# 示圖排版模式常數
IMG_MODE_AFTER_OPS = "after_ops"
IMG_MODE_LAST_PAGE = "last_page"

IMG_LAYOUT_OPTIONS = [
    (IMG_MODE_AFTER_OPS, "1. 工序後加圖片 (同頁下方/工序緊隨) 【標準版面】"),
    (IMG_MODE_LAST_PAGE, "2. 圖片統一置於最後一頁 (大圖附圖頁)")
]


def clean_program_name_helper(name):
    """程式檔名清洗輔助：移除結尾的 -Txx 或 _Txx，並清除空白"""
    if not name:
        return ""
    clean = str(name).strip()
    return re.sub(r'[-_]T\d+$', '', clean, flags=re.IGNORECASE).strip()


def get_item_program_name(op):
    """
    穩健解析工步所屬之程式檔名/群組名稱：
    優先順序：
    1. op.get("program_name")
    2. op.get("parent_program")
    3. op.get("group_name")
    4. op.get("clean_op_name")
    5. op.get("op_name")
    6. 若皆為空，預設為 "NC_PROGRAM"
    """
    if not isinstance(op, dict):
        return "NC_PROGRAM"

    cand = (
        op.get("program_name")
        or op.get("parent_program")
        or op.get("group_name")
        or op.get("clean_op_name")
        or op.get("op_name")
    )
    if cand:
        cand_str = str(cand).strip()
        cleaned = clean_program_name_helper(cand_str)
        if cleaned:
            return cleaned
        return cand_str
    return "NC_PROGRAM"


class InteractivePaginationDialogApp:
    def __init__(self, master, config):
        self.master = master
        self.config = config

        self.stages = config.get("stages", ["通用工段"])
        if not self.stages:
            self.stages = ["通用工段"]

        self.seq_style_var = tk.StringVar(value=config.get("initial_sequence_style", SEQ_STYLE_PER_PAGE))
        self.img_layout_var = tk.StringVar(value=config.get("initial_image_layout", IMG_MODE_AFTER_OPS))

        # 備份初始推薦資料供 [重設] 使用
        self.raw_default_pages = copy.deepcopy(config.get("default_pages", []))
        self.pages_data = []

        self.current_page_idx = 0
        self.drag_data = {"group_idx": None, "source_page": None}
        self._is_internal_selecting = False

        # 單元格內聯編輯器控制變數
        self._inline_entry = None
        self._inline_target_grp = None

        # 選中群組之操作綁定變數 (預設 need_capture 為 False，不勾選)
        self.curr_selected_grp_name_var = tk.StringVar(value="(未選取)")
        self.curr_selected_custom_seq_var = tk.StringVar(value="")
        self.curr_selected_capture_var = tk.BooleanVar(value=False)

        self.init_pages_from_config(self.raw_default_pages)
        self.setup_ui()
        self.refresh_all()

    def init_pages_from_config(self, raw_pages):
        """將輸入的 default_pages 初始化為內部頁面資料模型 (預設截圖為 False，不勾選)"""
        self.pages_data = []
        if not raw_pages:
            self.pages_data.append({
                "page_id": 1,
                "stage": self.stages[0],
                "user_stage_override": False,
                "items": []
            })
            return

        for p_idx, p in enumerate(raw_pages, start=1):
            stg = p.get("stage", self.stages[0])
            items = []
            raw_ops = p.get("operations", [])
            if not raw_ops and "rows" in p:
                for r in p["rows"]:
                    if isinstance(r, dict) and r.get("type") == "op":
                        items.append(copy.deepcopy(r.get("data", {})))
                    elif isinstance(r, dict) and "tool_number" in r:
                        items.append(copy.deepcopy(r))
            else:
                for op in raw_ops:
                    items.append(copy.deepcopy(op))

            # 確保每個工步都有健全的程式檔名標記與群組截圖開關 (預設 False)
            for op in items:
                prog = get_item_program_name(op)
                op["program_name"] = prog
                if not op.get("parent_program"):
                    op["parent_program"] = prog
                if not op.get("group_name"):
                    op["group_name"] = prog
                if "custom_seq" not in op:
                    op["custom_seq"] = ""
                if "need_capture" not in op:
                    op["need_capture"] = False  # 依使用者規則：預設為不勾選

            self.pages_data.append({
                "page_id": p_idx,
                "stage": stg,
                "user_stage_override": False,
                "items": items
            })

    def setup_ui(self):
        self.master.title("工單可視化互動式分頁排版編輯器 - 檔名群組與示圖排版模式")
        self.master.attributes("-topmost", True)
        self.master.minsize(1040, 700)
        self.master.configure(bg="#F4F6F9")

        # 居中計算
        win_w, win_h = 1080, 740
        scr_w = self.master.winfo_screenwidth()
        scr_h = self.master.winfo_screenheight()
        pos_x = max(20, (scr_w - win_w) // 2)
        pos_y = max(20, (scr_h - win_h) // 2)
        self.master.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

        # 1. 頂部深藍標題列 (Siemens NX 風格)
        header_frame = tk.Frame(self.master, bg="#004B87", padx=16, pady=8)
        header_frame.pack(fill=tk.X)

        lbl_title = tk.Label(
            header_frame,
            text="📑 工單可視化互動式分頁排版編輯器 (檔名群組連動 & 示圖排版決策)",
            font=("Microsoft JhengHei", 12, "bold"),
            fg="white",
            bg="#004B87"
        )
        lbl_title.pack(anchor="w")

        lbl_subtitle = tk.Label(
            header_frame,
            text="以「程式檔名為群組」整體排序與跨頁；可直接雙擊序號欄輸入文字標記工段 (不加序號)；群組截圖預設不勾選",
            font=("Microsoft JhengHei", 9),
            fg="#DCE6F1",
            bg="#004B87"
        )
        lbl_subtitle.pack(anchor="w", pady=(2, 0))

        # 2. 控制工具列 (分頁、重設、序號風格、示圖排版模式)
        toolbar = tk.Frame(self.master, bg="#E9ECEF", bd=1, relief=tk.SOLID, padx=12, pady=6)
        toolbar.pack(fill=tk.X)

        btn_add_page = tk.Button(
            toolbar,
            text=" ➕ 新增分頁 ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#28A745",
            fg="white",
            activebackground="#218838",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=2,
            command=self.on_add_page
        )
        btn_add_page.pack(side=tk.LEFT, padx=(0, 6))

        btn_reset = tk.Button(
            toolbar,
            text=" 🔄 重設初始推薦 ",
            font=("Microsoft JhengHei", 9),
            bg="#6C757D",
            fg="white",
            activebackground="#5A6268",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=2,
            command=self.on_reset_layout
        )
        btn_reset.pack(side=tk.LEFT, padx=(0, 12))

        # 序號風格選擇
        lbl_seq_style = tk.Label(
            toolbar,
            text="預設序號風格：",
            font=("Microsoft JhengHei", 9, "bold"),
            fg="#333333",
            bg="#E9ECEF"
        )
        lbl_seq_style.pack(side=tk.LEFT)

        self.cbo_seq_style = ttk.Combobox(
            toolbar,
            textvariable=self.seq_style_var,
            values=[opt[1] for opt in SEQ_STYLE_OPTIONS],
            state="readonly",
            width=28,
            font=("Microsoft JhengHei", 9)
        )
        for val_id, val_label in SEQ_STYLE_OPTIONS:
            if val_id == self.seq_style_var.get():
                self.cbo_seq_style.set(val_label)
                break
        self.cbo_seq_style.pack(side=tk.LEFT, padx=(0, 12))
        self.cbo_seq_style.bind("<<ComboboxSelected>>", self.on_seq_style_changed)

        # 示圖排版模式選擇
        lbl_img_mode = tk.Label(
            toolbar,
            text="示圖排版模式：",
            font=("Microsoft JhengHei", 9, "bold"),
            fg="#004B87",
            bg="#E9ECEF"
        )
        lbl_img_mode.pack(side=tk.LEFT)

        self.cbo_img_mode = ttk.Combobox(
            toolbar,
            textvariable=self.img_layout_var,
            values=[opt[1] for opt in IMG_LAYOUT_OPTIONS],
            state="readonly",
            width=32,
            font=("Microsoft JhengHei", 9)
        )
        for val_id, val_label in IMG_LAYOUT_OPTIONS:
            if val_id == self.img_layout_var.get():
                self.cbo_img_mode.set(val_label)
                break
        self.cbo_img_mode.pack(side=tk.LEFT, padx=(0, 12))
        self.cbo_img_mode.bind("<<ComboboxSelected>>", self.on_img_layout_changed)

        # 右側總覽統計
        self.lbl_stats = tk.Label(
            toolbar,
            text="總頁數：0 頁 | 總工步：0 刀",
            font=("Microsoft JhengHei", 9, "bold"),
            fg="#004B87",
            bg="#E9ECEF"
        )
        self.lbl_stats.pack(side=tk.RIGHT)

        # 3. 主工作區
        main_pane = tk.PanedWindow(self.master, orient=tk.HORIZONTAL, bg="#D8DDE4", sashwidth=5)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=12, pady=(8, 6))

        # === 左側面版：頁面導航清單 ===
        left_frame = tk.Frame(main_pane, bg="#FFFFFF", bd=1, relief=tk.SOLID)
        main_pane.add(left_frame, minsize=240, width=270)

        left_top_lbl = tk.Label(
            left_frame,
            text="【工單分頁清單】",
            font=("Microsoft JhengHei", 9, "bold"),
            fg="#222222",
            bg="#F1F3F5",
            pady=6,
            bd=1,
            relief=tk.SOLID
        )
        left_top_lbl.pack(fill=tk.X)

        self.page_listbox = tk.Listbox(
            left_frame,
            font=("Microsoft JhengHei", 9),
            bg="white",
            fg="#222222",
            selectbackground="#004B87",
            selectforeground="white",
            bd=0,
            highlightthickness=0
        )
        self.page_listbox.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.page_listbox.bind("<<ListboxSelect>>", self.on_page_select)

        # 左側底部分頁微調按鈕
        left_btn_bar = tk.Frame(left_frame, bg="#FFFFFF", padx=6, pady=6)
        left_btn_bar.pack(fill=tk.X)

        btn_del_page = tk.Button(
            left_btn_bar,
            text="🗑️ 刪除此頁",
            font=("Microsoft JhengHei", 9),
            bg="#DC3545",
            fg="white",
            activebackground="#C82333",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_delete_page
        )
        btn_del_page.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        btn_move_page_up = tk.Button(
            left_btn_bar,
            text="⬆ 往前",
            font=("Microsoft JhengHei", 9),
            bg="#E9ECEF",
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            command=lambda: self.on_move_page_order(-1)
        )
        btn_move_page_up.pack(side=tk.LEFT, padx=2)

        btn_move_page_down = tk.Button(
            left_btn_bar,
            text="⬇ 往後",
            font=("Microsoft JhengHei", 9),
            bg="#E9ECEF",
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            command=lambda: self.on_move_page_order(1)
        )
        btn_move_page_down.pack(side=tk.LEFT, padx=2)

        # === 右側面版：當前頁面工步編輯區 ===
        right_frame = tk.Frame(main_pane, bg="#FFFFFF", bd=1, relief=tk.SOLID)
        main_pane.add(right_frame, minsize=620)

        page_prop_frame = tk.Frame(right_frame, bg="#F8F9FA", bd=1, relief=tk.SOLID, padx=10, pady=5)
        page_prop_frame.pack(fill=tk.X)

        self.lbl_curr_page_title = tk.Label(
            page_prop_frame,
            text="第 1 頁配置",
            font=("Microsoft JhengHei", 10, "bold"),
            fg="#004B87",
            bg="#F8F9FA"
        )
        self.lbl_curr_page_title.pack(side=tk.LEFT, padx=(0, 10))

        lbl_assoc_stage = tk.Label(
            page_prop_frame,
            text="關聯示圖與原點：",
            font=("Microsoft JhengHei", 9),
            bg="#F8F9FA",
            fg="#333333"
        )
        lbl_assoc_stage.pack(side=tk.LEFT)

        self.cbo_stage = ttk.Combobox(
            page_prop_frame,
            values=self.stages,
            state="readonly",
            width=14,
            font=("Microsoft JhengHei", 9)
        )
        self.cbo_stage.pack(side=tk.LEFT, padx=(0, 12))
        self.cbo_stage.bind("<<ComboboxSelected>>", self.on_stage_changed)

        self.lbl_layout_badge = tk.Label(
            page_prop_frame,
            text="🟢 標準版面 (10刀以內)",
            font=("Microsoft JhengHei", 9, "bold"),
            fg="#155724",
            bg="#D4EDDA",
            padx=8,
            pady=2,
            bd=1,
            relief=tk.SOLID
        )
        self.lbl_layout_badge.pack(side=tk.RIGHT)

        # 依使用者需求：徹底移除上方提示列以精簡版面，文字輸入功能 100% 直觀融入下方序號欄儲存格
        self.ent_custom_seq = None
        self.lbl_selected_group_name = None

        # 右側中部：工步表格 Treeview
        tree_container = tk.Frame(right_frame, bg="white")
        tree_container.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        columns = ("seq", "prog", "tool_no", "spec", "flute", "holder", "time", "stage", "capture", "note")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", selectmode="extended")

        self.tree.heading("seq", text="序號 / 標記 [可點擊輸入]")
        self.tree.heading("prog", text="程式檔名 / 工序群組")
        self.tree.heading("tool_no", text="刀號")
        self.tree.heading("spec", text="刀具名稱型號規格")
        self.tree.heading("flute", text="刃長")
        self.tree.heading("holder", text="夾長")
        self.tree.heading("time", text="時間")
        self.tree.heading("stage", text="原工段")
        self.tree.heading("capture", text="工序截圖")
        self.tree.heading("note", text="切削參數 / 安全警示")

        self.tree.column("seq", width=110, anchor="center")
        self.tree.column("prog", width=175, anchor="w")
        self.tree.column("tool_no", width=55, anchor="center")
        self.tree.column("spec", width=120, anchor="w")
        self.tree.column("flute", width=50, anchor="center")
        self.tree.column("holder", width=50, anchor="center")
        self.tree.column("time", width=70, anchor="center")
        self.tree.column("stage", width=65, anchor="center")
        self.tree.column("capture", width=65, anchor="center")
        self.tree.column("note", width=140, anchor="w")

        v_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        h_scroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        tree_container.rowconfigure(0, weight=1)
        tree_container.columnconfigure(0, weight=1)

        # 綁定選取、雙擊內聯編輯、單擊截圖切換與拖曳
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select_expand_group)
        self.tree.bind("<ButtonPress-1>", self.on_tree_button_press)
        self.tree.bind("<B1-Motion>", self.on_tree_mouse_drag)
        self.tree.bind("<ButtonRelease-1>", self.on_tree_button_release)
        self.tree.bind("<Double-Button-1>", self.on_tree_double_click)
        self.tree.bind("<F2>", self.on_tree_f2_pressed)

        # 右側底部：以「程式檔名為群組」之操作工具列
        item_action_bar = tk.Frame(right_frame, bg="#F1F3F5", bd=1, relief=tk.SOLID, padx=8, pady=5)
        item_action_bar.pack(fill=tk.X)

        btn_grp_up = tk.Button(
            item_action_bar,
            text=" ⬆ 群組上移 ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#FFFFFF",
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            command=lambda: self.on_move_group_order(-1)
        )
        btn_grp_up.pack(side=tk.LEFT, padx=(0, 4))

        btn_grp_down = tk.Button(
            item_action_bar,
            text=" ⬇ 群組下移 ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#FFFFFF",
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            command=lambda: self.on_move_group_order(1)
        )
        btn_grp_down.pack(side=tk.LEFT, padx=(0, 10))

        btn_to_prev = tk.Button(
            item_action_bar,
            text=" ⬅ 群組移至上一頁 ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#E3F2FD",
            fg="#0D47A1",
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            command=self.on_move_group_to_prev_page
        )
        btn_to_prev.pack(side=tk.LEFT, padx=(0, 4))

        btn_to_next = tk.Button(
            item_action_bar,
            text=" ➔ 群組移至下一頁 ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#E3F2FD",
            fg="#0D47A1",
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            command=self.on_move_group_to_next_page
        )
        btn_to_next.pack(side=tk.LEFT, padx=(0, 10))

        lbl_target_p = tk.Label(item_action_bar, text="移至指定頁：", font=("Microsoft JhengHei", 9), bg="#F1F3F5")
        lbl_target_p.pack(side=tk.LEFT)

        self.cbo_target_page = ttk.Combobox(item_action_bar, state="readonly", width=10, font=("Microsoft JhengHei", 9))
        self.cbo_target_page.pack(side=tk.LEFT, padx=(0, 4))

        btn_move_target = tk.Button(
            item_action_bar,
            text="移動群組",
            font=("Microsoft JhengHei", 9),
            bg="#FFFFFF",
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            command=self.on_move_group_to_target_page
        )
        btn_move_target.pack(side=tk.LEFT, padx=(0, 12))

        btn_isolate_page = tk.Button(
            item_action_bar,
            text=" 📄 選取群組獨立成新頁 ",
            font=("Microsoft JhengHei", 9, "bold"),
            bg="#FFF3CD",
            fg="#856404",
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            command=self.on_isolate_group_to_new_page
        )
        btn_isolate_page.pack(side=tk.RIGHT)

        # 4. 底部全域決策欄
        bottom_bar = tk.Frame(self.master, bg="#EAEEF3", padx=16, pady=8)
        bottom_bar.pack(fill=tk.X, side=tk.BOTTOM)

        lbl_tips = tk.Label(
            bottom_bar,
            text="💡 提示：可直接在「序號欄」輸入文字標記工段 (直接顯示不加序號)；群組截圖預設不勾選，點擊群組首行可單獨勾選。",
            font=("Microsoft JhengHei", 9),
            fg="#555555",
            bg="#EAEEF3"
        )
        lbl_tips.pack(side=tk.LEFT)

        self.btn_confirm = tk.Button(
            bottom_bar,
            text=" ✔ 確認排版並產出工單 (Enter) ",
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#0078D7",
            fg="white",
            activebackground="#005A9E",
            activeforeground="white",
            padx=14,
            pady=5,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_confirm
        )
        self.btn_confirm.pack(side=tk.RIGHT, padx=(8, 0))

        btn_cancel = tk.Button(
            bottom_bar,
            text=" ❌ 取消 ",
            font=("Microsoft JhengHei", 9),
            bg="#D0D5DD",
            fg="#333333",
            activebackground="#B0B5BD",
            activeforeground="black",
            padx=12,
            pady=5,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.on_cancel
        )
        btn_cancel.pack(side=tk.RIGHT)

        self.master.bind("<Return>", lambda e: self.on_confirm())
        self.master.bind("<Escape>", lambda e: self.on_cancel())
        self.master.protocol("WM_DELETE_WINDOW", self.on_cancel)

    # ==================== 頁面與群組核心演算法 ====================

    def get_selected_img_layout_mode(self):
        """解析當前選擇之示圖排版模式代碼"""
        selected_text = self.cbo_img_mode.get()
        for opt_id, opt_label in IMG_LAYOUT_OPTIONS:
            if opt_label == selected_text:
                return opt_id
        return IMG_MODE_AFTER_OPS

    def on_img_layout_changed(self, event=None):
        self.refresh_all()

    def get_page_groups(self, page_idx=None):
        """
        將指定頁面的工步按相鄰且相同程式檔名切分為群組 (Groups)。
        每個群組擁有 custom_seq 與 need_capture 屬性 (預設 need_capture=False)。
        """
        if page_idx is None:
            page_idx = self.current_page_idx
        if not (0 <= page_idx < len(self.pages_data)):
            return []

        items = self.pages_data[page_idx]["items"]
        if not items:
            return []

        groups = []
        curr_prog = None
        curr_items = []
        curr_indices = []

        for i, op in enumerate(items):
            p_name = get_item_program_name(op)
            if curr_prog is None:
                curr_prog = p_name
                curr_items = [op]
                curr_indices = [i]
            elif p_name == curr_prog:
                curr_items.append(op)
                curr_indices.append(i)
            else:
                c_seq = curr_items[0].get("custom_seq", "")
                c_cap = curr_items[0].get("need_capture", False)
                groups.append({
                    "group_idx": len(groups),
                    "prog_name": curr_prog,
                    "custom_seq": c_seq,
                    "need_capture": c_cap,
                    "items": curr_items,
                    "indices": curr_indices,
                    "iids": [f"item_{idx}" for idx in curr_indices]
                })
                curr_prog = p_name
                curr_items = [op]
                curr_indices = [i]

        if curr_items:
            c_seq = curr_items[0].get("custom_seq", "")
            c_cap = curr_items[0].get("need_capture", False)
            groups.append({
                "group_idx": len(groups),
                "prog_name": curr_prog,
                "custom_seq": c_seq,
                "need_capture": c_cap,
                "items": curr_items,
                "indices": curr_indices,
                "iids": [f"item_{idx}" for idx in curr_indices]
            })

        return groups

    def get_selected_seq_style(self):
        selected_text = self.cbo_seq_style.get()
        for opt_id, opt_label in SEQ_STYLE_OPTIONS:
            if opt_label == selected_text:
                return opt_id
        return SEQ_STYLE_PER_PAGE

    def on_seq_style_changed(self, event=None):
        self.refresh_tree_items()

    def refresh_all(self):
        """全域刷新頁面清單、下拉選單與當前頁面內容"""
        for i, p in enumerate(self.pages_data, start=1):
            p["page_id"] = i

        total_ops = sum(len(p["items"]) for p in self.pages_data)
        self.lbl_stats.config(text=f"總頁數：{len(self.pages_data)} 頁 | 總工步：{total_ops} 刀")

        img_mode = self.get_selected_img_layout_mode()

        self.page_listbox.delete(0, tk.END)
        for i, p in enumerate(self.pages_data, start=1):
            num = len(p["items"])
            stg = p["stage"]
            # 檢查該頁是否有勾選截圖的群組
            has_capture = any(op.get("need_capture", False) for op in p["items"])

            if num == 0:
                tag = "⚪ 空白"
            elif img_mode == IMG_MODE_LAST_PAGE:
                tag = "📋 刀具頁"
            else:  # after_ops
                if not has_capture:
                    tag = "📋 無示圖"
                elif num <= 10:
                    tag = "🟢 標準"
                elif num <= 15:
                    tag = "🟡 延伸"
                elif num <= 20:
                    tag = "🟠 微縮"
                else:
                    tag = "🔵 大圖頁"
            self.page_listbox.insert(tk.END, f"第 {i} 頁 ({num}刀) [{stg}] {tag}")

        if self.current_page_idx >= len(self.pages_data):
            self.current_page_idx = max(0, len(self.pages_data) - 1)

        if len(self.pages_data) > 0:
            self.page_listbox.selection_clear(0, tk.END)
            self.page_listbox.selection_set(self.current_page_idx)
            self.page_listbox.see(self.current_page_idx)

        page_options = [f"第 {i} 頁" for i in range(1, len(self.pages_data) + 1)]
        self.cbo_target_page.config(values=page_options)
        if page_options:
            self.cbo_target_page.set(page_options[0])

        self.refresh_active_page_details()

    def refresh_active_page_details(self):
        """刷新右側當前頁面之屬性與項目表格"""
        if not self.pages_data or self.current_page_idx >= len(self.pages_data):
            return

        page = self.pages_data[self.current_page_idx]
        p_id = page["page_id"]
        items = page["items"]
        num_items = len(items)
        img_mode = self.get_selected_img_layout_mode()

        # 智慧多數工段判定
        if not page.get("user_stage_override", False) and items:
            stage_counts = Counter(op.get("stage", "") for op in items if op.get("stage"))
            if stage_counts:
                majority_stage = stage_counts.most_common(1)[0][0]
                if majority_stage in self.stages:
                    page["stage"] = majority_stage

        self.lbl_curr_page_title.config(text=f"第 {p_id} 頁配置 (共 {num_items} 刀)")
        self.cbo_stage.set(page["stage"])

        # 版面徽章依排版模式呈現
        has_capture = any(op.get("need_capture", False) for op in items)
        if num_items == 0:
            self.lbl_layout_badge.config(text="⚪ 空白頁面 (0 刀)", fg="#6C757D", bg="#E2E3E5")
        elif img_mode == IMG_MODE_LAST_PAGE:
            self.lbl_layout_badge.config(text=f"📋 純刀具清單頁 ({num_items} 刀，示圖統一置於最後)", fg="#004085", bg="#CCE5FF")
        else:
            if not has_capture:
                self.lbl_layout_badge.config(text=f"📋 無示圖頁面 ({num_items} 刀，均未勾選截圖)", fg="#495057", bg="#E9ECEF")
            elif num_items <= 10:
                self.lbl_layout_badge.config(text=f"🟢 標準版面 ({num_items}/10 刀，包含示圖)", fg="#155724", bg="#D4EDDA")
            elif num_items <= 15:
                self.lbl_layout_badge.config(text=f"🟡 單頁延伸 ({num_items} 刀，自動增格維持示圖)", fg="#856404", bg="#FFF3CD")
            elif num_items <= 20:
                self.lbl_layout_badge.config(text=f"🟠 單頁延伸 ({num_items} 刀，示圖同頁微縮)", fg="#D9534F", bg="#FDF7E7")
            else:
                self.lbl_layout_badge.config(text=f"🔵 刀具延伸頁 ({num_items} 刀，示圖自動移至下一頁大圖頁)", fg="#004085", bg="#CCE5FF")

        self.refresh_tree_items()

    def refresh_tree_items(self):
        """重新填入 Treeview 並格式化檔名群組、自訂序號與截圖狀態標記"""
        self.cancel_inline_edit()
        self.tree.delete(*self.tree.get_children())
        if not self.pages_data or self.current_page_idx >= len(self.pages_data):
            return

        page = self.pages_data[self.current_page_idx]
        items = page["items"]
        page_stage = page["stage"]
        seq_style = self.get_selected_seq_style()

        # 計算全域累積工步數
        global_offset = 0
        if seq_style == SEQ_STYLE_GLOBAL:
            for p_i in range(self.current_page_idx):
                global_offset += len(self.pages_data[p_i]["items"])

        groups = self.get_page_groups(self.current_page_idx)

        # 遍歷各群組並插入表格
        curr_row_idx = 0
        for grp in groups:
            p_name = grp["prog_name"]
            grp_custom_seq = grp.get("custom_seq", "")
            grp_capture = grp.get("need_capture", False)
            grp_items = grp["items"]

            for sub_i, op in enumerate(grp_items):
                curr_row_idx += 1
                idx = curr_row_idx

                # 需求 1：若使用者有輸入自訂工段文字，序號完全依照輸入文字顯示，不須再加流水序號
                if grp_custom_seq:
                    raw_seq = grp_custom_seq
                else:
                    if seq_style == SEQ_STYLE_STAGE_PREFIX:
                        raw_seq = f"{page_stage}-{idx:02d}"
                    elif seq_style == SEQ_STYLE_GLOBAL:
                        raw_seq = str(global_offset + idx)
                    elif seq_style == SEQ_STYLE_ZERO_PAD:
                        raw_seq = f"{idx:02d}"
                    elif seq_style == SEQ_STYLE_STEP_10:
                        raw_seq = str(idx * 10)
                    else:
                        raw_seq = str(idx)

                # 依使用者最新決策 (方案 A)：序號儲存格以輸入框樣式【 [ 1 ] 】顯示，直觀引導使用者此處具備輸入功能
                seq_str = f"[ {raw_seq} ]"

                # 程式檔名視覺呈現 (首行縮排)
                if sub_i == 0:
                    if grp_custom_seq:
                        display_prog = f"📁 [{grp_custom_seq}] {p_name}"
                    else:
                        display_prog = f"📁 {p_name}"
                else:
                    display_prog = f"   ↳ (同上: {p_name})"

                # 需求 2：截圖勾選欄只對「程式群組」生效，僅在首行顯示，後續刀具為空
                if sub_i == 0:
                    capture_display = "☑ 截圖 (📷)" if grp_capture else "☐ 略過"
                else:
                    capture_display = ""

                tool_no = op.get("tool_number", "-")
                spec = op.get("tool_spec_display", op.get("tool_name", "-"))
                flute = str(op.get("flute_length", "-"))
                holder = str(op.get("holder_length", "-"))
                time_str = str(op.get("time", "-"))
                stage_str = str(op.get("stage", page_stage))
                if grp_custom_seq:
                    stage_str = grp_custom_seq

                note_parts = []
                if op.get("has_tool_conflict"):
                    note_parts.append(op.get("tool_conflict_msg", "【刀號衝突】"))
                if op.get("tool_safety_warning"):
                    note_parts.append(op.get("tool_safety_warning", ""))
                if op.get("note"):
                    note_parts.append(str(op.get("note")))
                note_str = " ".join(note_parts).strip()

                item_id = f"item_{curr_row_idx - 1}"
                tags = ()
                if op.get("has_tool_conflict"):
                    tags = ("conflict",)

                self.tree.insert(
                    "",
                    tk.END,
                    iid=item_id,
                    values=(seq_str, display_prog, tool_no, spec, flute, holder, time_str, stage_str, capture_display, note_str),
                    tags=tags
                )

        self.tree.tag_configure("conflict", background="#FFEBEB", foreground="#A00000")

    # ==================== 單元格內聯編輯 (In-place Cell Editing on Seq) ====================

    def start_inline_seq_edit(self, item_id):
        """
        核心互動功能 1：
        直接在「序號 / 標記」儲存格上建立浮動 Entry 供使用者直接輸入文字！
        按 Enter 立即套用至該群組旗下所有刀具，不加流水序號。
        """
        if not item_id or not self.tree.exists(item_id):
            return

        self.cancel_inline_edit()

        row_idx = int(item_id.split("_")[1])
        groups = self.get_page_groups(self.current_page_idx)
        target_grp = None
        for grp in groups:
            if row_idx in grp["indices"]:
                target_grp = grp
                break

        if not target_grp:
            return

        curr_val = target_grp.get("custom_seq", "")
        if not curr_val:
            try:
                raw_cell = self.tree.item(item_id, "values")[0]
                curr_val = str(raw_cell).strip("[] ")
            except Exception:
                curr_val = ""

        bbox = self.tree.bbox(item_id, "seq")

        if not bbox:
            # 備援容錯：若 bbox 暫不可用 (例如在非可見視窗測試)，帶入變數
            self.curr_selected_custom_seq_var.set(curr_val)
            if hasattr(self, "ent_custom_seq") and self.ent_custom_seq:
                try:
                    self.ent_custom_seq.focus_set()
                    self.ent_custom_seq.select_range(0, tk.END)
                except Exception:
                    pass
            return

        x, y, w, h = bbox
        self._inline_target_grp = target_grp

        self._inline_entry = tk.Entry(
            self.tree,
            font=("Microsoft JhengHei", 9, "bold"),
            justify="center",
            bd=2,
            relief=tk.SOLID,
            bg="#FFFFD0"  # 醒目淡黃底色
        )
        self._inline_entry.place(x=x + 2, y=y + 1, width=w - 4, height=h - 2)
        if curr_val:
            self._inline_entry.insert(0, curr_val)
        self._inline_entry.select_range(0, tk.END)
        self._inline_entry.focus_set()

        def commit_edit(event=None):
            if not self._inline_entry:
                return "break"
            new_text = self._inline_entry.get().strip()
            # 若輸入中包含外圍括號 [ ]，自動去除以獲取純文字
            if new_text.startswith("[") and new_text.endswith("]"):
                new_text = new_text[1:-1].strip()
            self.cancel_inline_edit()
            # 套用至該群組旗下所有工步 (若清空則設為 ""，自動恢復預設流水號)
            for op in target_grp["items"]:
                op["custom_seq"] = new_text
                if new_text:
                    op["stage"] = new_text
            self.refresh_active_page_details()
            return "break"

        def cancel_edit(event=None):
            self.cancel_inline_edit()
            return "break"

        self._inline_entry.bind("<Return>", commit_edit)
        self._inline_entry.bind("<Escape>", cancel_edit)
        self._inline_entry.bind("<FocusOut>", commit_edit)

    def cancel_inline_edit(self):
        """銷毀內聯輸入框"""
        if self._inline_entry:
            try:
                self._inline_entry.destroy()
            except Exception:
                pass
            self._inline_entry = None
            self._inline_target_grp = None

    def on_tree_double_click(self, event):
        """雙擊表格事件：雙擊序號欄位 (#1)，立即啟動單元格內聯編輯"""
        col = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        if col == "#1" and row:
            self.start_inline_seq_edit(row)

    def on_tree_f2_pressed(self, event=None):
        """按下 F2 快捷鍵：直接啟動選中項目之序號內聯編輯"""
        sel = self.tree.selection()
        if sel:
            self.start_inline_seq_edit(sel[0])

    # ==================== 選取連動與自訂標記/截圖控制 ====================

    def on_tree_select_expand_group(self, event=None):
        """
        核心聯動選取機制：
        1. 點選任何單一刀具時，自動全選所屬整個程式檔名群組的所有刀具。
        2. 同步將該群組的名稱、自訂工段文字標記與截圖勾選狀態載入至操作欄位。
        """
        if self._is_internal_selecting:
            return

        selected_iids = self.tree.selection()
        if not selected_iids:
            self.curr_selected_grp_name_var.set("(未選取)")
            self.curr_selected_custom_seq_var.set("")
            self.curr_selected_capture_var.set(False)
            return

        groups = self.get_page_groups(self.current_page_idx)
        selected_indices = set()
        for iid in selected_iids:
            if iid.startswith("item_"):
                try:
                    selected_indices.add(int(iid.split("_")[1]))
                except Exception:
                    pass

        target_iids = set()
        first_selected_grp = None
        for grp in groups:
            if any(idx in selected_indices for idx in grp["indices"]):
                target_iids.update(grp["iids"])
                if first_selected_grp is None:
                    first_selected_grp = grp

        # 自動全選補齊
        if target_iids and set(selected_iids) != target_iids:
            self._is_internal_selecting = True
            try:
                self.tree.selection_set(*list(target_iids))
            finally:
                self._is_internal_selecting = False

        # 更新群組專屬操作欄位
        if first_selected_grp:
            self.curr_selected_grp_name_var.set(f"📁 {first_selected_grp['prog_name']}")
            self.curr_selected_custom_seq_var.set(first_selected_grp.get("custom_seq", ""))
            self.curr_selected_capture_var.set(first_selected_grp.get("need_capture", False))

    def on_apply_custom_seq(self):
        """將上方文字框輸入的工段標記套用至目前選取群組"""
        sel_g_indices = self.get_selected_group_indices()
        if not sel_g_indices:
            messagebox.showinfo("提示", "請先在表格中點選要標記工段的程式群組！")
            return

        input_text = self.curr_selected_custom_seq_var.get().strip()
        groups = self.get_page_groups(self.current_page_idx)

        for g_idx in sel_g_indices:
            if 0 <= g_idx < len(groups):
                for op in groups[g_idx]["items"]:
                    op["custom_seq"] = input_text
                    if input_text:
                        op["stage"] = input_text

        self.refresh_active_page_details()
        self.restore_group_selection(sel_g_indices)

    def on_clear_custom_seq(self):
        """清除目前選取群組的自訂工段文字標記，恢復為風格流水號"""
        self.curr_selected_custom_seq_var.set("")
        self.on_apply_custom_seq()

    def on_toggle_group_capture(self):
        """切換目前選取群組的截圖開關 (need_capture)"""
        sel_g_indices = self.get_selected_group_indices()
        if not sel_g_indices:
            return

        is_checked = self.curr_selected_capture_var.get()
        groups = self.get_page_groups(self.current_page_idx)

        for g_idx in sel_g_indices:
            if 0 <= g_idx < len(groups):
                for op in groups[g_idx]["items"]:
                    op["need_capture"] = is_checked

        self.refresh_active_page_details()
        self.restore_group_selection(sel_g_indices)

    def restore_group_selection(self, group_indices):
        """輔助函式：操作後恢復選取指定群組"""
        updated_groups = self.get_page_groups(self.current_page_idx)
        to_select_iids = []
        for g_i in group_indices:
            if 0 <= g_i < len(updated_groups):
                to_select_iids.extend(updated_groups[g_i]["iids"])
        if to_select_iids:
            self._is_internal_selecting = True
            try:
                self.tree.selection_set(*to_select_iids)
            finally:
                self._is_internal_selecting = False

    def get_selected_group_indices(self):
        """取得當前選取之所有群組在 groups 列表中的索引 (已排序)"""
        selected_iids = self.tree.selection()
        if not selected_iids:
            return []

        selected_row_indices = set()
        for iid in selected_iids:
            if iid.startswith("item_"):
                try:
                    selected_row_indices.add(int(iid.split("_")[1]))
                except Exception:
                    pass

        groups = self.get_page_groups(self.current_page_idx)
        matched_group_indices = []
        for g_idx, grp in enumerate(groups):
            if any(idx in selected_row_indices for idx in grp["indices"]):
                matched_group_indices.append(g_idx)

        matched_group_indices.sort()
        return matched_group_indices

    # ==================== 群組排序與移動操作 ====================

    def on_move_group_order(self, direction):
        """在當前頁面內調整程式群組順序 (-1 群組上移, 1 群組下移)"""
        groups = self.get_page_groups(self.current_page_idx)
        sel_g_indices = self.get_selected_group_indices()
        if not sel_g_indices:
            messagebox.showinfo("提示", "請先在表格中選取要調整順序的程式群組！")
            return

        new_groups = list(groups)
        if direction < 0:
            for g_idx in sel_g_indices:
                if g_idx > 0 and (g_idx - 1) not in sel_g_indices:
                    new_groups[g_idx], new_groups[g_idx - 1] = new_groups[g_idx - 1], new_groups[g_idx]
            new_sel = [max(0, g - 1) for g in sel_g_indices]
        else:
            for g_idx in reversed(sel_g_indices):
                if g_idx < len(new_groups) - 1 and (g_idx + 1) not in sel_g_indices:
                    new_groups[g_idx], new_groups[g_idx + 1] = new_groups[g_idx + 1], new_groups[g_idx]
            new_sel = [min(len(new_groups) - 1, g + 1) for g in sel_g_indices]

        flattened_items = []
        for grp in new_groups:
            flattened_items.extend(grp["items"])
        self.pages_data[self.current_page_idx]["items"] = flattened_items

        self.refresh_active_page_details()
        self.restore_group_selection(new_sel)

    def move_selected_groups_to_page(self, target_page_idx):
        """將當前頁選取的整個程式群組完整移至目標頁"""
        if target_page_idx == self.current_page_idx or not (0 <= target_page_idx < len(self.pages_data)):
            return

        groups = self.get_page_groups(self.current_page_idx)
        sel_g_indices = self.get_selected_group_indices()
        if not sel_g_indices:
            messagebox.showinfo("提示", "請先選取要移動的程式群組！")
            return

        moving_items = []
        remaining_items = []

        for g_idx, grp in enumerate(groups):
            if g_idx in sel_g_indices:
                moving_items.extend(grp["items"])
            else:
                remaining_items.extend(grp["items"])

        self.pages_data[self.current_page_idx]["items"] = remaining_items
        self.pages_data[target_page_idx]["items"].extend(moving_items)

        self.refresh_all()

    def on_move_group_to_prev_page(self):
        if self.current_page_idx > 0:
            self.move_selected_groups_to_page(self.current_page_idx - 1)
        else:
            messagebox.showinfo("提示", "目前已是第一頁，無法再往前移！")

    def on_move_group_to_next_page(self):
        if self.current_page_idx < len(self.pages_data) - 1:
            self.move_selected_groups_to_page(self.current_page_idx + 1)
        else:
            sel_g_indices = self.get_selected_group_indices()
            if not sel_g_indices:
                messagebox.showinfo("提示", "請先選擇要移動的程式群組！")
                return
            resp = messagebox.askyesno("新增分頁確認", "目前已是最後一頁，是否要新增下一頁並將選取的程式群組移入？")
            if resp:
                new_p_idx = len(self.pages_data)
                self.pages_data.append({
                    "page_id": new_p_idx + 1,
                    "stage": self.pages_data[self.current_page_idx]["stage"],
                    "user_stage_override": False,
                    "items": []
                })
                self.move_selected_groups_to_page(new_p_idx)
                self.current_page_idx = new_p_idx
                self.refresh_all()
    def on_move_group_to_target_page(self):
        target_str = self.cbo_target_page.get()
        m = re.search(r'\d+', target_str)
        if m:
            target_p = int(m.group()) - 1
            self.move_selected_groups_to_page(target_p)

    def on_isolate_group_to_new_page(self):
        """選取之整個程式群組獨立建立全新分頁"""
        groups = self.get_page_groups(self.current_page_idx)
        sel_g_indices = self.get_selected_group_indices()
        if not sel_g_indices:
            messagebox.showinfo("提示", "請先選取要獨立成頁的程式群組！")
            return

        moving_items = []
        remaining_items = []
        for g_idx, grp in enumerate(groups):
            if g_idx in sel_g_indices:
                moving_items.extend(grp["items"])
            else:
                remaining_items.extend(grp["items"])

        self.pages_data[self.current_page_idx]["items"] = remaining_items

        new_page_stage = moving_items[0].get("stage", self.stages[0]) if moving_items else self.stages[0]
        self.pages_data.insert(self.current_page_idx + 1, {
            "page_id": len(self.pages_data) + 1,
            "stage": new_page_stage,
            "user_stage_override": False,
            "items": moving_items
        })
        self.current_page_idx += 1
        self.refresh_all()

    # ==================== 滑鼠點擊與拖曳 (Click & Drag & Drop) ====================

    def on_tree_button_press(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)

        # 需求 1：單擊序號欄位 (#1)，直接啟動儲存格輸入框！
        if col == "#1" and item and item.startswith("item_"):
            self.start_inline_seq_edit(item)
            return "break"

        # 需求 2：單擊截圖欄位 (#9)，僅在點擊群組首行 (可見勾選標記處) 時切換該程式群組的截圖開關
        if col == "#9" and item and item.startswith("item_"):
            try:
                row_idx = int(item.split("_")[1])
                groups = self.get_page_groups(self.current_page_idx)
                for grp in groups:
                    # 嚴格限制：僅點擊在該群組首行 (首個刀具) 才切換，後續空白列不觸發切換
                    if row_idx == grp["indices"][0]:
                        new_state = not grp.get("need_capture", False)
                        for op in grp["items"]:
                            op["need_capture"] = new_state
                        self.refresh_active_page_details()
                        return "break"
            except Exception:
                pass

        if item and item.startswith("item_"):
            try:
                row_idx = int(item.split("_")[1])
                groups = self.get_page_groups(self.current_page_idx)
                for g_idx, grp in enumerate(groups):
                    if row_idx in grp["indices"]:
                        self.drag_data["group_idx"] = g_idx
                        self.drag_data["source_page"] = self.current_page_idx
                        break
            except Exception:
                pass

    def on_tree_mouse_drag(self, event):
        if self.drag_data.get("group_idx") is not None:
            self.master.config(cursor="fleur")

    def on_tree_button_release(self, event):
        self.master.config(cursor="")
        dragged_g_idx = self.drag_data.get("group_idx")
        self.drag_data["group_idx"] = None
        if dragged_g_idx is None:
            return

        # 1. 檢查是否釋放在左側的分頁清單上 (跨頁拖曳)
        listbox_x = self.page_listbox.winfo_rootx()
        listbox_y = self.page_listbox.winfo_rooty()
        listbox_w = self.page_listbox.winfo_width()
        listbox_h = self.page_listbox.winfo_height()

        mouse_rx = self.master.winfo_pointerx()
        mouse_ry = self.master.winfo_pointery()

        if listbox_x <= mouse_rx <= listbox_x + listbox_w and listbox_y <= mouse_ry <= listbox_y + listbox_h:
            rel_y = mouse_ry - listbox_y
            target_list_idx = self.page_listbox.nearest(rel_y)
            if 0 <= target_list_idx < len(self.pages_data) and target_list_idx != self.current_page_idx:
                groups = self.get_page_groups(self.current_page_idx)
                if 0 <= dragged_g_idx < len(groups):
                    self._is_internal_selecting = True
                    try:
                        self.tree.selection_set(*groups[dragged_g_idx]["iids"])
                    finally:
                        self._is_internal_selecting = False
                    self.move_selected_groups_to_page(target_list_idx)
                return

        # 2. 同頁內拖曳調整群組順序
        target_iid = self.tree.identify_row(event.y)
        if target_iid and target_iid.startswith("item_"):
            try:
                target_row = int(target_iid.split("_")[1])
                groups = self.get_page_groups(self.current_page_idx)
                target_g_idx = None
                for g_idx, grp in enumerate(groups):
                    if target_row in grp["indices"]:
                        target_g_idx = g_idx
                        break

                if target_g_idx is not None and target_g_idx != dragged_g_idx:
                    new_groups = list(groups)
                    grp_elem = new_groups.pop(dragged_g_idx)
                    new_groups.insert(target_g_idx, grp_elem)

                    flattened = []
                    for g in new_groups:
                        flattened.extend(g["items"])
                    self.pages_data[self.current_page_idx]["items"] = flattened
                    self.refresh_active_page_details()
            except Exception:
                pass

    # ==================== 分頁增刪與風格切換 ====================

    def on_page_select(self, event=None):
        sel = self.page_listbox.curselection()
        if sel:
            self.current_page_idx = sel[0]
            self.refresh_active_page_details()

    def on_stage_changed(self, event=None):
        if self.current_page_idx < len(self.pages_data):
            page = self.pages_data[self.current_page_idx]
            page["stage"] = self.cbo_stage.get()
            page["user_stage_override"] = True
            self.refresh_all()

    def on_add_page(self):
        new_id = len(self.pages_data) + 1
        default_stage = self.stages[0]
        if self.pages_data:
            default_stage = self.pages_data[-1]["stage"]

        self.pages_data.append({
            "page_id": new_id,
            "stage": default_stage,
            "user_stage_override": False,
            "items": []
        })
        self.current_page_idx = len(self.pages_data) - 1
        self.refresh_all()

    def on_delete_page(self):
        if len(self.pages_data) <= 1:
            messagebox.showwarning("提示", "工單必須至少保留 1 個分頁！")
            return

        curr_page = self.pages_data[self.current_page_idx]
        items_to_rescue = curr_page["items"]

        if items_to_rescue:
            resp = messagebox.askyesno(
                "刪除確認",
                f"第 {curr_page['page_id']} 頁目前包含 {len(items_to_rescue)} 道工步。\n確定要刪除此頁，並將這些程式群組自動移入鄰近頁面嗎？"
            )
            if not resp:
                return

        target_idx = self.current_page_idx - 1 if self.current_page_idx > 0 else 1
        self.pages_data[target_idx]["items"].extend(items_to_rescue)
        del self.pages_data[self.current_page_idx]

        self.current_page_idx = min(self.current_page_idx, len(self.pages_data) - 1)
        self.refresh_all()

    def on_move_page_order(self, direction):
        target = self.current_page_idx + direction
        if 0 <= target < len(self.pages_data):
            self.pages_data[self.current_page_idx], self.pages_data[target] = (
                self.pages_data[target],
                self.pages_data[self.current_page_idx]
            )
            self.current_page_idx = target
            self.refresh_all()

    def on_reset_layout(self):
        resp = messagebox.askyesno("重設確認", "確定要放棄當前的分頁調整，恢復為系統推薦的初始排版嗎？")
        if resp:
            self.init_pages_from_config(self.raw_default_pages)
            self.current_page_idx = 0
            self.refresh_all()

    # ==================== 輸出建構與對話框回傳 ====================

    def build_export_pages(self):
        """
        將互動排版成果轉換為符合 export_multipage_via_vbs 規格之 pages 列表：
        - 序號規格：若有 custom_seq，保留至各工步物件；op_name 依群組邊界首行填寫、後續留空
        - 截圖規則：工序右方截圖若有勾選，即表示該工序需要一張專屬截圖 (以工序 prog_name 為單位，不再合併工段)
        - 示圖排版模式：
          A. 工序後加圖片 (after_ops)：
             - 頁面內若無勾選截圖之工序：純刀具頁面 (最多30刀，show_image=False)
             - 頁面內若有 1 個勾選工序：
               - <=10 刀：同頁下方展示 (image_top_row=18, show_image=True, target_op=prog_name)
               - 11~20 刀：單頁延伸下方展示 (image_top_row=8+N, show_image=True, target_op=prog_name)
               - >20 刀：純刀具頁 + 緊隨生成一頁大圖附圖頁 (target_op=prog_name)
             - 頁面內若有多個勾選工序：
               - 刀具列作為純刀具頁，緊隨其後為每個勾選工序各生成一頁專屬大圖頁 (target_op=prog_name)
          B. 圖片統一置於最後一頁 (last_page)：
             - 前面所有頁面全為純刀具頁面 (每頁最多30刀，show_image=False)
             - 最後為全工單中所有勾選截圖之工序，各生成一頁專屬大圖附圖頁 (target_op=prog_name, is_appendix_image_page=True)
        """
        export_pages = []
        img_layout_mode = self.get_selected_img_layout_mode()

        # 收集全工單中所有需要截圖的工序群組清單 (維持出現順序，去重複)
        captured_groups = []
        seen_prog_names = set()
        for p_idx, page in enumerate(self.pages_data):
            groups = self.get_page_groups(p_idx)
            for grp in groups:
                if grp.get("need_capture", False):
                    p_name = grp["prog_name"]
                    if p_name not in seen_prog_names:
                        seen_prog_names.add(p_name)
                        captured_groups.append({
                            "prog_name": p_name,
                            "stage": grp["items"][0].get("stage", page["stage"]) if grp["items"] else page["stage"],
                            "custom_seq": grp.get("custom_seq", "")
                        })

        if img_layout_mode == IMG_MODE_LAST_PAGE:
            # === 模式 B：圖片統一置於最後一頁 (大圖附圖頁) ===
            # 1. 前面所有頁面全為純刀具頁面 (show_image = False)
            for p_idx, page in enumerate(self.pages_data, start=1):
                page_stage = page["stage"]
                groups = self.get_page_groups(p_idx - 1)
                formatted_ops = []
                for grp in groups:
                    p_name = grp["prog_name"]
                    c_seq = grp.get("custom_seq", "")
                    c_cap = grp.get("need_capture", False)
                    for sub_i, op in enumerate(grp["items"]):
                        op_copy = copy.deepcopy(op)
                        op_copy["program_name"] = p_name
                        op_copy["custom_seq"] = c_seq
                        op_copy["need_capture"] = c_cap
                        if sub_i == 0:
                            op_copy["op_name"] = p_name
                        else:
                            op_copy["op_name"] = ""
                        formatted_ops.append(op_copy)

                rem = formatted_ops
                if not rem:
                    export_pages.append({
                        "stage": page_stage,
                        "rows": [],
                        "show_image": False,
                        "image_top_row": None,
                        "image_bottom_row": None,
                        "is_appendix_image_page": False
                    })
                while rem:
                    batch = rem[:30]
                    rem = rem[30:]
                    export_pages.append({
                        "stage": page_stage,
                        "rows": [{"type": "op", "data": op} for op in batch],
                        "show_image": False,
                        "image_top_row": None,
                        "image_bottom_row": None,
                        "is_appendix_image_page": False
                    })

            # 2. 於所有刀具頁面之後，盡可能將所有勾選截圖之工序置於同一附圖頁 (每頁最多 6 張示圖，排不下再分頁)
            MAX_IMAGES_PER_APPENDIX_PAGE = 6
            for chunk_start in range(0, len(captured_groups), MAX_IMAGES_PER_APPENDIX_PAGE):
                chunk = captured_groups[chunk_start:chunk_start + MAX_IMAGES_PER_APPENDIX_PAGE]
                chunk_tgt_ops = [c["prog_name"] for c in chunk]
                if len(chunk) == 1:
                    layout = "single"
                elif len(chunk) == 2:
                    layout = "side_by_side"
                else:
                    layout = "grid"

                export_pages.append({
                    "stage": chunk[0]["stage"],
                    "target_ops": chunk_tgt_ops,
                    "target_op": chunk[0]["prog_name"],  # 向後相容單圖欄位
                    "rows": [],
                    "show_image": True,
                    "image_top_row": 7,
                    "image_bottom_row": 36,
                    "is_appendix_image_page": True,
                    "image_layout": layout
                })

        else:
            # === 模式 A：工序後加圖片 (選項 C：盡可能於排版可接受的範圍內(工序+截圖) 排成同一頁，排不下的再建立分頁) ===
            for p_idx, page in enumerate(self.pages_data, start=1):
                page_stage = page["stage"]
                groups = self.get_page_groups(p_idx - 1)

                formatted_groups = []
                for grp in groups:
                    p_name = grp["prog_name"]
                    c_seq = grp.get("custom_seq", "")
                    c_cap = grp.get("need_capture", False)
                    grp_items = []
                    for sub_i, op in enumerate(grp["items"]):
                        op_copy = copy.deepcopy(op)
                        op_copy["program_name"] = p_name
                        op_copy["custom_seq"] = c_seq
                        op_copy["need_capture"] = c_cap
                        op_copy["op_name"] = p_name if sub_i == 0 else ""
                        grp_items.append(op_copy)
                    formatted_groups.append({
                        "prog_name": p_name,
                        "stage": grp["items"][0].get("stage", page_stage) if grp["items"] else page_stage,
                        "need_capture": c_cap,
                        "items": grp_items
                    })

                if not formatted_groups:
                    export_pages.append({
                        "stage": page_stage,
                        "rows": [],
                        "show_image": False,
                        "image_top_row": None,
                        "image_bottom_row": None,
                        "is_appendix_image_page": False
                    })
                else:
                    def flush_chunk(chunk):
                        if not chunk:
                            return
                        chunk_ops = [op for g in chunk for op in g["items"]]
                        chunk_cap_ops = [g["prog_name"] for g in chunk if g["need_capture"]]
                        n_ops = len(chunk_ops)
                        n_caps = len(chunk_cap_ops)
                        stg = chunk[0]["stage"] if chunk else page_stage

                        if n_caps == 0:
                            export_pages.append({
                                "stage": stg,
                                "rows": [{"type": "op", "data": op} for op in chunk_ops],
                                "show_image": False,
                                "image_top_row": None,
                                "image_bottom_row": None,
                                "is_appendix_image_page": False
                            })
                        elif n_caps == 1:
                            export_pages.append({
                                "stage": stg,
                                "target_op": chunk_cap_ops[0],
                                "target_ops": [chunk_cap_ops[0]],
                                "image_layout": "single",
                                "rows": [{"type": "op", "data": op} for op in chunk_ops],
                                "show_image": True,
                                "image_top_row": 18 if n_ops <= 10 else (8 + n_ops),
                                "image_bottom_row": 36,
                                "is_appendix_image_page": False
                            })
                        else:  # n_caps == 2 (雙工序左右並排展示)
                            export_pages.append({
                                "stage": stg,
                                "target_op": chunk_cap_ops[0],
                                "target_ops": [chunk_cap_ops[0], chunk_cap_ops[1]],
                                "image_layout": "side_by_side",
                                "rows": [{"type": "op", "data": op} for op in chunk_ops],
                                "show_image": True,
                                "image_top_row": 18 if n_ops <= 10 else (8 + n_ops),
                                "image_bottom_row": 36,
                                "is_appendix_image_page": False
                            })

                    current_chunk = []

                    for grp in formatted_groups:
                        n_grp_ops = len(grp["items"])
                        g_cap = grp["need_capture"]

                        # 單一群組若本身超大 (> 20 刀且需拍照，或 > 30 刀)
                        if (g_cap and n_grp_ops > 20) or n_grp_ops > 30:
                            if current_chunk:
                                flush_chunk(current_chunk)
                                current_chunk = []

                            rem = grp["items"]
                            while rem:
                                batch = rem[:30]
                                rem = rem[30:]
                                export_pages.append({
                                    "stage": grp["stage"],
                                    "rows": [{"type": "op", "data": op} for op in batch],
                                    "show_image": False,
                                    "image_top_row": None,
                                    "image_bottom_row": None,
                                    "is_appendix_image_page": False
                                })
                            if g_cap:
                                export_pages.append({
                                    "stage": grp["stage"],
                                    "target_op": grp["prog_name"],
                                    "target_ops": [grp["prog_name"]],
                                    "image_layout": "single",
                                    "rows": [],
                                    "show_image": True,
                                    "image_top_row": 7,
                                    "image_bottom_row": 36,
                                    "is_appendix_image_page": True
                                })
                            continue

                        # 嘗試將 grp 加入 current_chunk
                        tentative = current_chunk + [grp]
                        tentative_ops = sum(len(g["items"]) for g in tentative)
                        tentative_caps = sum(1 for g in tentative if g["need_capture"])

                        can_fit = False
                        if tentative_caps == 0:
                            can_fit = (tentative_ops <= 30)
                        elif tentative_caps in (1, 2):
                            can_fit = (tentative_ops <= 20)
                        else:  # tentative_caps > 2 (單頁最多容納 2 張並排圖片)
                            can_fit = False

                        if can_fit:
                            current_chunk = tentative
                        else:
                            # 當前頁排不下了，結算當前頁，另立新分頁
                            if current_chunk:
                                flush_chunk(current_chunk)
                            current_chunk = [grp]

                    if current_chunk:
                        flush_chunk(current_chunk)
        return export_pages

    def on_confirm(self):
        valid_pages = [p for p in self.pages_data if p["items"]]
        if not valid_pages and self.pages_data:
            valid_pages = [self.pages_data[0]]
        self.pages_data = valid_pages

        final_pages = self.build_export_pages()
        seq_style = self.get_selected_seq_style()
        img_mode = self.get_selected_img_layout_mode()

        self.result = {
            "action": "ok",
            "sequence_style": seq_style,
            "image_layout_mode": img_mode,
            "pages": final_pages
        }
        self.master.destroy()

    def on_cancel(self):
        self.result = {
            "action": "cancel",
            "sequence_style": SEQ_STYLE_PER_PAGE,
            "image_layout_mode": IMG_MODE_AFTER_OPS,
            "pages": []
        }
        self.master.destroy()


def parse_args():
    parser = argparse.ArgumentParser(description="工單可視化互動式分頁排版編輯器")
    parser.add_argument("--cfg-file", type=str, default="", help="輸入配置檔 (JSON 格式)")
    parser.add_argument("--res-file", type=str, default="", help="輸出結果檔 (JSON 格式)")
    parser.add_argument("--auto-action", type=str, default="", help="單元測試自動動作 (ok, cancel)")
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
    app = InteractivePaginationDialogApp(root, config)

    if args.auto_action:
        def trigger_auto():
            if args.auto_action == "ok":
                app.on_confirm()
            else:
                app.on_cancel()
        root.after(100, trigger_auto)

    root.mainloop()

    if args.res_file:
        res_data = getattr(app, "result", {"action": "cancel", "sequence_style": "per_page", "image_layout_mode": "after_ops", "pages": []})
        tmp_res = args.res_file + ".tmp"
        try:
            with open(tmp_res, "w", encoding="utf-8") as f:
                json.dump(res_data, f, ensure_ascii=False, indent=2)
                f.flush()
                os.fsync(f.fileno())
            if os.path.exists(args.res_file):
                try:
                    os.remove(args.res_file)
                except Exception:
                    pass
            os.rename(tmp_res, args.res_file)
        except Exception as ex:
            try:
                if os.path.exists(tmp_res):
                    os.remove(tmp_res)
            except Exception:
                pass


if __name__ == "__main__":
    main()
