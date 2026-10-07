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
import argparse
import tkinter as tk
import ctypes

user32 = ctypes.windll.user32

def parse_args():
    parser = argparse.ArgumentParser(description="CNC 加工示圖精緻拍照小工具")
    parser.add_argument("--stage", type=str, default="通用工段", help="當前工段名稱")
    parser.add_argument("--index", type=int, default=1, help="當前工段序號")
    parser.add_argument("--total", type=int, default=1, help="總工段數")
    parser.add_argument("--parent-hwnd", type=int, default=0, help="NX 主視窗句柄")
    parser.add_argument("--req-file", type=str, default="", help="擺正請求旗標檔案路徑")
    return parser.parse_args()

def main():
    args = parse_args()
    stage_name = args.stage
    curr_idx = args.index
    total_stages = args.total
    parent_hwnd = args.parent_hwnd
    req_file = args.req_file

    # 若未指定 parent-hwnd，嘗試抓取當前前景視窗 (即呼叫前的 NX 主視窗)
    if parent_hwnd == 0:
        parent_hwnd = user32.GetForegroundWindow()

    display_name = stage_name if stage_name != "通用工段" else "加工全貌"

    root = tk.Tk()
    root.title(f"📷 拍照引導 ({curr_idx}/{total_stages})")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    root.configure(bg="#F0F4F8")

    # 精緻小尺寸，置於螢幕右上角 (避免遮擋工作區)
    win_w = 360
    win_h = 135
    screen_w = root.winfo_screenwidth()
    pos_x = max(10, screen_w - win_w - 40)
    pos_y = 60
    root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

    exit_code = [2] # 預設 2 為略過或取消

    def trigger_f8_snap(event=None):
        """觸發 NX 畫面擺正 (F8 貼齊正交視圖)"""
        # 1. 寫入跨進程旗標，直接由 NX 主進程底層執行原生視圖擺正 (100% 絕對生效)
        if req_file:
            try:
                with open(req_file, "w") as f:
                    f.write("snap")
            except Exception:
                pass

        # 2. 備援：嘗試將焦點歸還給 NX 主視窗並模擬發送按鍵
        if parent_hwnd and user32.IsWindow(parent_hwnd):
            try:
                user32.SetForegroundWindow(parent_hwnd)
            except Exception:
                pass
        try:
            user32.keybd_event(0x77, 0, 0, 0)
            user32.keybd_event(0x77, 0, 2, 0)
        except Exception:
            pass

    def on_confirm():
        exit_code[0] = 0 # 0 代表確認拍照
        root.destroy()

    def on_skip():
        exit_code[0] = 2 # 2 代表略過
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_skip)

    # 視窗全域監聽鍵盤 F8 鍵
    root.bind("<F8>", trigger_f8_snap)

    # 標題與工段標籤
    lbl_title = tk.Label(
        root, text=f"📷 工段示圖 ({curr_idx}/{total_stages})：{display_name}",
        font=("Microsoft JhengHei", 10, "bold"), bg="#F0F4F8", fg="#004B87"
    )
    lbl_title.pack(pady=(10, 2))

    lbl_hint = tk.Label(
        root, text="中鍵旋轉視圖 | 支援鍵盤 F8 或點擊下方 45° 擺正",
        font=("Microsoft JhengHei", 9), bg="#F0F4F8", fg="#555555"
    )
    lbl_hint.pack(pady=(0, 8))

    # 按鈕列
    btn_frame = tk.Frame(root, bg="#F0F4F8")
    btn_frame.pack()

    # 按鈕 1：畫面擺正 (F8)
    btn_snap = tk.Button(
        btn_frame, text=" 📐 45° 擺正 (F8) ", command=trigger_f8_snap,
        bg="#008080", fg="white", activebackground="#006666", activeforeground="white",
        font=("Microsoft JhengHei", 9, "bold"), padx=8, pady=4, relief=tk.FLAT, cursor="hand2"
    )
    btn_snap.pack(side=tk.LEFT, padx=4)

    # 按鈕 2：立即拍照
    btn_confirm = tk.Button(
        btn_frame, text=" 📸 立即拍照 ", command=on_confirm,
        bg="#0078D7", fg="white", activebackground="#005A9E", activeforeground="white",
        font=("Microsoft JhengHei", 9, "bold"), padx=10, pady=4, relief=tk.FLAT, cursor="hand2"
    )
    btn_confirm.pack(side=tk.LEFT, padx=4)

    # 按鈕 3：略過此段
    btn_skip = tk.Button(
        btn_frame, text=" 略過此段 ", command=on_skip,
        bg="#E1E5EA", fg="#444444", activebackground="#D0D5DD", activeforeground="#000000",
        font=("Microsoft JhengHei", 9), padx=8, pady=4, relief=tk.FLAT, cursor="hand2"
    )
    btn_skip.pack(side=tk.LEFT, padx=4)

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
