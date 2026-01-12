# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
側邊導航欄組件
"""

import customtkinter as ctk
from typing import Optional, Callable, Any, List

from ui.theme import theme_manager, Fonts


class SidebarItem(ctk.CTkFrame):
    """側邊欄導航項目"""
    
    def __init__(
        self,
        master: Any,
        text: str,
        icon: str = "",
        command: Optional[Callable] = None,
        **kwargs
    ):
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self._text = text
        self._icon = icon
        self._command = command
        self._selected = False
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # 整個項目作為可點擊區域
        self.configure(height=40, cursor="hand2")
        self.pack_propagate(False)
        
        # 內部容器
        self.inner = ctk.CTkFrame(self, fg_color="transparent", corner_radius=8)
        self.inner.pack(fill="both", expand=True, padx=8, pady=2)
        
        # 圖示 (固定寬度)
        self.lbl_icon = ctk.CTkLabel(
            self.inner,
            text=self._icon,
            font=("Segoe UI Emoji", 14),
            width=28,
            anchor="center"
        )
        self.lbl_icon.pack(side="left", padx=(12, 0))
        
        # 文字
        self.lbl_text = ctk.CTkLabel(
            self.inner,
            text=self._text,
            font=Fonts.to_tuple(Fonts.BODY),
            text_color="#37474f",
            anchor="w"
        )
        self.lbl_text.pack(side="left", padx=(8, 12))
        
        # 綁定點擊事件到所有子元件
        for widget in [self, self.inner, self.lbl_icon, self.lbl_text]:
            widget.bind("<Button-1>", lambda e: self._on_click())
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
    
    def _on_enter(self, event=None):
        """滑鼠進入"""
        if not self._selected:
            self.inner.configure(fg_color="#e1f0ff")
    
    def _on_leave(self, event=None):
        """滑鼠離開"""
        if not self._selected:
            self.inner.configure(fg_color="transparent")
    
    def _on_click(self):
        """點擊事件"""
        if self._command:
            self._command()
    
    def set_selected(self, selected: bool):
        """設定選中狀態"""
        self._selected = selected
        if selected:
            self.inner.configure(fg_color="#e1f0ff")
            self.lbl_text.configure(text_color="#1e88e5")
        else:
            self.inner.configure(fg_color="transparent")
            self.lbl_text.configure(text_color="#37474f")


class Sidebar(ctk.CTkFrame):
    """側邊導航欄"""
    
    def __init__(
        self,
        master: Any,
        title: str = "WIM 管理工具",
        version: str = "3.0",
        on_navigate: Optional[Callable[[str], None]] = None,
        **kwargs
    ):
        super().__init__(
            master,
            fg_color="#f8fbff",
            corner_radius=0,
            width=180,
            **kwargs
        )
        self.pack_propagate(False)
        
        self._title = title
        self._version = version
        self._on_navigate = on_navigate
        self._items: dict[str, SidebarItem] = {}
        self._current_page = None
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # === 導航項目 ===
        self.nav_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.nav_frame.pack(fill="both", expand=True, pady=(16, 0))
        
        # === 底部開發者資訊 ===
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", side="bottom", pady=(0, 12))
        
        # 分隔線
        ctk.CTkFrame(
            footer,
            height=1,
            fg_color="#e1f0ff"
        ).pack(fill="x", padx=16, pady=(0, 10))
        
        # 開發者資訊（一行式）
        ctk.CTkLabel(
            footer,
            text="Developed by RexPeng",
            font=("Microsoft JhengHei UI", 11),
            text_color="#90a4ae"
        ).pack(anchor="center")
    
    def add_item(self, key: str, text: str, icon: str = ""):
        """新增導航項目"""
        item = SidebarItem(
            self.nav_frame,
            text=text,
            icon=icon,
            command=lambda k=key: self._navigate_to(k)
        )
        item.pack(fill="x")
        self._items[key] = item
        
        # 第一個項目預設選中（只更新狀態，不觸發回調）
        if len(self._items) == 1:
            item.set_selected(True)
            self._current_page = key
    
    def _navigate_to(self, key: str):
        """導航到指定頁面"""
        if self._current_page == key:
            return
        
        # 更新選中狀態
        for k, item in self._items.items():
            item.set_selected(k == key)
        
        self._current_page = key
        
        if self._on_navigate:
            self._on_navigate(key)
    
    def get_current_page(self) -> Optional[str]:
        """取得當前頁面"""
        return self._current_page
