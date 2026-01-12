# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自定義對話框組件
與主程序同樣式的 UI 對話框
"""

import customtkinter as ctk
from typing import Optional, Callable
from ui.theme import Fonts, theme_manager


class ModernDialog(ctk.CTkToplevel):
    """現代風格對話框"""
    
    def __init__(
        self,
        parent,
        title: str = "提示",
        message: str = "",
        detail: str = "",
        icon: str = "info",  # info, warning, error, question
        buttons: list = None,  # [("按鈕文字", "類型", 回調), ...]
        width: int = 400,
        height: int = None,  # 自動計算
        **kwargs
    ):
        super().__init__(parent, **kwargs)
        
        self._result = None
        self._buttons_config = buttons or [("確定", "primary", None)]
        
        # 視窗設定
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # 設定外觀
        self.configure(fg_color="#ffffff")
        
        # 計算高度
        if height is None:
            base_height = 120
            if detail:
                base_height += 40 + len(detail) // 40 * 20
            height = min(base_height + len(message) // 50 * 20, 300)
        
        # 視窗大小與位置
        self.geometry(f"{width}x{height}")
        self._center_window(parent, width, height)
        
        # 建立 UI
        self._build_ui(title, message, detail, icon)
        
        # 綁定關閉事件
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Escape>", lambda e: self._on_close())
        
        # 等待視窗關閉
        self.wait_window()
    
    def _center_window(self, parent, width: int, height: int):
        """視窗置中"""
        parent.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        
        x = px + (pw - width) // 2
        y = py + (ph - height) // 2
        
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def _build_ui(self, title: str, message: str, detail: str, icon: str):
        """建立 UI"""
        # 主容器
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=16)
        
        # 圖示與標題區
        header = ctk.CTkFrame(main_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))
        
        # 圖示
        icon_colors = {
            "info": "#64b5f6",
            "warning": "#ffb74d", 
            "error": "#ef5350",
            "question": "#9575cd"
        }
        icon_symbols = {
            "info": "ℹ",
            "warning": "⚠",
            "error": "✕",
            "question": "?"
        }
        
        icon_label = ctk.CTkLabel(
            header,
            text=icon_symbols.get(icon, "ℹ"),
            font=("Microsoft JhengHei UI", 24, "bold"),
            text_color=icon_colors.get(icon, "#64b5f6"),
            width=40
        )
        icon_label.pack(side="left")
        
        # 標題
        ctk.CTkLabel(
            header,
            text=title,
            font=Fonts.to_tuple(Fonts.TITLE_SMALL),
            text_color="#37474f"
        ).pack(side="left", padx=(8, 0))
        
        # 訊息
        msg_label = ctk.CTkLabel(
            main_frame,
            text=message,
            font=Fonts.to_tuple(Fonts.BODY),
            text_color="#546e7a",
            wraplength=350,
            justify="left",
            anchor="w"
        )
        msg_label.pack(fill="x", pady=(0, 8))
        
        # 詳細說明（如果有）
        if detail:
            detail_frame = ctk.CTkFrame(
                main_frame,
                fg_color="#f5f5f5",
                corner_radius=6
            )
            detail_frame.pack(fill="x", pady=(0, 12))
            
            ctk.CTkLabel(
                detail_frame,
                text=detail,
                font=Fonts.to_tuple(Fonts.BODY_SMALL),
                text_color="#78909c",
                wraplength=340,
                justify="left",
                anchor="w"
            ).pack(padx=12, pady=10)
        
        # 按鈕區
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(8, 0))
        
        # 按鈕樣式
        button_styles = {
            "primary": {
                "fg_color": "#64b5f6",
                "hover_color": "#42a5f5",
                "text_color": "#ffffff"
            },
            "secondary": {
                "fg_color": "#e0e0e0",
                "hover_color": "#bdbdbd",
                "text_color": "#37474f"
            },
            "danger": {
                "fg_color": "#ef5350",
                "hover_color": "#e53935",
                "text_color": "#ffffff"
            },
            "success": {
                "fg_color": "#81c784",
                "hover_color": "#66bb6a",
                "text_color": "#ffffff"
            }
        }
        
        # 建立按鈕（從右到左排列）
        for i, btn_config in enumerate(reversed(self._buttons_config)):
            text = btn_config[0]
            style = btn_config[1] if len(btn_config) > 1 else "primary"
            callback = btn_config[2] if len(btn_config) > 2 else None
            
            style_config = button_styles.get(style, button_styles["primary"])
            
            btn = ctk.CTkButton(
                btn_frame,
                text=text,
                font=Fonts.to_tuple(Fonts.BODY),
                width=100,
                height=36,
                corner_radius=6,
                **style_config,
                command=lambda cb=callback, t=text: self._on_button_click(t, cb)
            )
            btn.pack(side="right", padx=(0, 12 if i > 0 else 0))
    
    def _on_button_click(self, text: str, callback: Optional[Callable]):
        """按鈕點擊"""
        self._result = text
        if callback:
            callback()
        self.destroy()
    
    def _on_close(self):
        """關閉對話框"""
        self._result = None
        self.destroy()
    
    @property
    def result(self):
        """取得結果"""
        return self._result


class ConfirmDialog(ModernDialog):
    """確認對話框"""
    
    def __init__(
        self,
        parent,
        title: str = "確認",
        message: str = "",
        detail: str = "",
        confirm_text: str = "確定",
        cancel_text: str = "取消",
        icon: str = "question",
        **kwargs
    ):
        super().__init__(
            parent,
            title=title,
            message=message,
            detail=detail,
            icon=icon,
            buttons=[
                (cancel_text, "secondary", None),
                (confirm_text, "primary", None)
            ],
            **kwargs
        )
    
    @property
    def confirmed(self) -> bool:
        """是否確認"""
        return self._result == self._buttons_config[1][0]


class UnmountWarningDialog(ModernDialog):
    """卸載警告對話框"""
    
    def __init__(self, parent, mount_dir: str, **kwargs):
        message = "在卸載前，請確認以下事項："
        detail = (
            "• 請關閉所有開啟掛載目錄的資料夾視窗\n"
            "• 請關閉所有正在存取掛載目錄的程式\n"
            "• 請確認沒有檔案正在複製或寫入中\n\n"
            f"掛載目錄：{mount_dir}\n\n"
            "如果卸載失敗，請嘗試使用「修復卸載」按鈕。"
        )
        
        super().__init__(
            parent,
            title="卸載提醒",
            message=message,
            detail=detail,
            icon="warning",
            buttons=[
                ("取消", "secondary", None),
                ("確定卸載", "primary", None)
            ],
            width=450,
            height=300,
            **kwargs
        )
    
    @property
    def confirmed(self) -> bool:
        """是否確認卸載"""
        return self._result == "確定卸載"
