# -*- coding: utf-8 -*-
"""
組件名稱：machine_profile_manager.py
功能職責：CNC 機台配置與檔頭樣板管理器
設計原則：
1. 組件功能單純原則：負責機型設定清單維護、設定檔持久化 (JSON) 與記憶上次選擇。
2. 零破壞與多階防禦 (Zero-Debug Protocol)：設定檔損毀或缺失時自動回退預設值。
3. 支援機型：
   - FANUC_HORIZONTAL: Fanuc 臥式機型 (含 G30、M11/M10 B0.、M06 T 專屬置換)
   - FANUC_VERTICAL: Fanuc 立式機型 (保留原後處理碼 + 注入標準工單註解)
   - BROTHER: BROTHER 兄弟機 (保留原後處理碼 + 注入標準工單註解)
   - RAW: 原始後處理不變
"""

import os
import json

SETTINGS_FILE_NAME = "nc_machine_settings.json"

DEFAULT_PROFILES = {
    "FANUC_HORIZONTAL": {
        "id": "FANUC_HORIZONTAL",
        "name": "Fanuc 臥式機 (B0/G30/四軸專用檔頭)",
        "mode": "REPLACE_HORIZONTAL_HEADER",
        "comment_char": ["(", ")"],
        "has_percent": True,
        "default_post_name": "Fanuc",
        "header_template": [
            "G91G30Z0.",
            "G91G30X0.Y0.",
            "G90G17G00G40G49G80",
            "M11",
            "G90G0G54B0.",
            "M10",
            "M06 T{tool_no}",
            "S{speed} M03",
            "G17 G40 G80",
            "G90 G54 G00 X{start_x} Y{start_y}",
            "G0 G43 H{tool_no} Z{safe_z}",
            "M08"
        ]
    },
    "FANUC_VERTICAL": {
        "id": "FANUC_VERTICAL",
        "name": "Fanuc 立式機 (保留原程式碼 + 注入工單註解)",
        "mode": "INJECT_HEADER_COMMENTS",
        "comment_char": ["(", ")"],
        "has_percent": True,
        "default_post_name": "Fanuc",
        "header_template": []
    },
    "BROTHER": {
        "id": "BROTHER",
        "name": "BROTHER 機台 (高速攻牙機專用)",
        "mode": "INJECT_HEADER_COMMENTS",
        "comment_char": ["(", ")"],
        "has_percent": True,
        "default_post_name": "Fanuc",
        "header_template": []
    },
    "RAW": {
        "id": "RAW",
        "name": "原始後處理輸出 (不變更檔頭)",
        "mode": "RAW_OUTPUT",
        "comment_char": ["(", ")"],
        "has_percent": True,
        "default_post_name": "Fanuc",
        "header_template": []
    }
}

def get_settings_path():
    cur_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(cur_dir, SETTINGS_FILE_NAME)

def load_machine_profiles():
    """
    載入所有可用機台配置清單。若本地有覆寫則合併載入。
    """
    cfg_path = get_settings_path()
    profiles = dict(DEFAULT_PROFILES)

    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                custom_profiles = data.get("profiles", {})
                for k, v in custom_profiles.items():
                    if k in profiles:
                        profiles[k].update(v)
                    else:
                        profiles[k] = v
        except Exception:
            pass
    return profiles

def get_post_settings():
    """
    讀取上次工程師選取的機台與後處理器設定，包含機型檔頭 ID、後處理器名稱與自訂路徑。
    回傳字典：
    {
        "last_selected_id": "FANUC_HORIZONTAL",
        "last_selected_post": "Fanuc_2026",
        "last_custom_post_path": ""
    }
    """
    cfg_path = get_settings_path()
    res = {
        "last_selected_id": "FANUC_HORIZONTAL",
        "last_selected_post": "Fanuc_2026",
        "last_custom_post_path": ""
    }
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    last_id = data.get("last_selected_id")
                    if last_id and last_id in DEFAULT_PROFILES:
                        res["last_selected_id"] = last_id
                    last_post = data.get("last_selected_post")
                    if last_post and isinstance(last_post, str):
                        res["last_selected_post"] = last_post
                    last_cust = data.get("last_custom_post_path")
                    if last_cust and isinstance(last_cust, str):
                        res["last_custom_post_path"] = last_cust
        except Exception:
            pass
    return res

def save_post_settings(profile_id=None, post_name=None, custom_post_path=None):
    """
    保存本次工程師選取的機台設定 (機型檔頭、後處理器名稱、自訂後處理路徑)。
    採用 Zero-Debug 原子寫入 (.tmp + fsync + replace)，防止斷電或衝突損毀。
    """
    cfg_path = get_settings_path()
    data = {}
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if not isinstance(data, dict):
                    data = {}
        except Exception:
            data = {}

    if profile_id is not None:
        data["last_selected_id"] = str(profile_id)
    if post_name is not None:
        data["last_selected_post"] = str(post_name)
    if custom_post_path is not None:
        data["last_custom_post_path"] = str(custom_post_path)

    # Zero-Debug 原子寫入
    tmp_path = cfg_path + ".tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        try:
            os.replace(tmp_path, cfg_path)
        except OSError:
            if os.path.exists(cfg_path):
                os.remove(cfg_path)
            os.rename(tmp_path, cfg_path)
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass

def get_last_selected_profile_id():
    """
    讀取上次工程師選取的機台 ID，預設為 FANUC_HORIZONTAL。
    """
    return get_post_settings().get("last_selected_id", "FANUC_HORIZONTAL")

def save_last_selected_profile_id(profile_id):
    """
    保存本次工程師選取的機台 ID。
    """
    save_post_settings(profile_id=profile_id)
