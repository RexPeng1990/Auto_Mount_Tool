# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
可折疊區塊組件
"""

import customtkinter as ctk
from typing import Optional, Any, Callable

from ui.theme import theme_manager, Fonts, Spacing


class CollapsibleSection(ctk.CTkFrame):
    """
    可折疊的區塊組件
    用於建立 Accordion 風格的介面
    """
    
    def __init__(
        self,
        master: Any,
        title: str,
        icon: str = "",
        default_expanded: bool = True,
        header_buttons: Optional[list] = None,
        on_toggle: Optional[Callable[[bool], None]] = None,
        **kwargs
    ):
        """
        初始化可折疊區塊
        
        Args:
            master: 父組件
            title: 區塊標題
            icon: 圖示 (emoji)
            default_expanded: 預設是否展開
            header_buttons: 標題列右側按鈕 [(text, command), ...]
            on_toggle: 展開/收合回調
        """
        super().__init__(
            master,
            fg_color=theme_manager.colors.card_bg,
            corner_radius=8,
            border_width=1,
            border_color=theme_manager.colors.border,
            **kwargs
        )
        
        self._title = title
        self._icon = icon
        self._expanded = default_expanded
        self._header_buttons = header_buttons or []
        self._on_toggle = on_toggle
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # === 標題列 ===
        self.header = ctk.CTkFrame(self, fg_color="transparent", height=32)
        self.header.pack(fill="x", padx=10, pady=(6, 0))
        self.header.pack_propagate(False)
        
        # 左側：箭頭 + 圖示 + 標題（可點擊）
        left_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        left_frame.pack(side="left", fill="y")
        left_frame.bind("<Button-1>", lambda e: self.toggle())
        
        self.arrow_label = ctk.CTkLabel(
            left_frame,
            text="▼" if self._expanded else "▶",
            font=Fonts.to_tuple(Fonts.BODY_SMALL),
            text_color=theme_manager.colors.text_muted,
            width=18
        )
        self.arrow_label.pack(side="left")
        self.arrow_label.bind("<Button-1>", lambda e: self.toggle())
        
        title_text = f"{self._icon} {self._title}" if self._icon else self._title
        self.title_label = ctk.CTkLabel(
            left_frame,
            text=title_text,
            font=Fonts.to_tuple(Fonts.TITLE_SMALL),
            text_color=theme_manager.colors.text_primary
        )
        self.title_label.pack(side="left", padx=(4, 0))
        self.title_label.bind("<Button-1>", lambda e: self.toggle())
        
        # 右側：額外按鈕
        if self._header_buttons:
            btn_frame = ctk.CTkFrame(self.header, fg_color="transparent")
            btn_frame.pack(side="right")
            
            for btn_text, btn_command in self._header_buttons:
                btn = ctk.CTkButton(
                    btn_frame,
                    text=btn_text,
                    width=60,
                    height=26,
                    font=Fonts.to_tuple(Fonts.CAPTION),
                    fg_color="transparent",
                    hover_color=theme_manager.colors.border,
                    text_color=theme_manager.colors.text_muted,
                    command=btn_command
                )
                btn.pack(side="left", padx=(4, 0))
        
        # === 內容區域 ===
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        if self._expanded:
            self.content_frame.pack(fill="both", expand=True, padx=10, pady=(6, 8))
    
    def toggle(self):
        """切換展開/收合狀態"""
        self._expanded = not self._expanded
        
        if self._expanded:
            self.content_frame.pack(fill="both", expand=True, padx=10, pady=(6, 8))
            self.arrow_label.configure(text="▼")
        else:
            self.content_frame.pack_forget()
            self.arrow_label.configure(text="▶")
        
        if self._on_toggle:
            self._on_toggle(self._expanded)
    
    def expand(self):
        """展開"""
        if not self._expanded:
            self.toggle()
    
    def collapse(self):
        """收合"""
        if self._expanded:
            self.toggle()
    
    def is_expanded(self) -> bool:
        """是否展開"""
        return self._expanded
    
    def get_content_frame(self) -> ctk.CTkFrame:
        """取得內容框架，用於添加子組件"""
        return self.content_frame
    
    def set_title(self, title: str):
        """設定標題"""
        self._title = title
        title_text = f"{self._icon} {title}" if self._icon else title
        self.title_label.configure(text=title_text)
