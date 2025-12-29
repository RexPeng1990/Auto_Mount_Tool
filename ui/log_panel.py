# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日誌面板組件
現代化日誌顯示
"""

import customtkinter as ctk
import tkinter as tk
from typing import Optional, Any
from datetime import datetime
import re

from ui.theme import ThemeManager, Fonts, Spacing, theme_manager


class LogPanel(ctk.CTkFrame):
    """
    現代化日誌面板
    支援彩色輸出、時間戳、過濾
    """
    
    # 日誌級別對應顏色
    LEVEL_COLORS = {
        "INFO": "#3b82f6",      # 藍色
        "SUCCESS": "#22c55e",   # 綠色
        "WARNING": "#f59e0b",   # 黃色
        "ERROR": "#ef4444",     # 紅色
        "DEBUG": "#6b7280",     # 灰色
    }
    
    # 自動檢測模式
    PATTERNS = {
        r"^✓|成功|完成|done|success": "SUCCESS",
        r"^✗|失敗|錯誤|error|fail": "ERROR",
        r"警告|注意|warning": "WARNING",
        r"^\[DEBUG\]|debug": "DEBUG",
    }
    
    def __init__(
        self,
        master: Any,
        title: str = "系統日誌",
        show_timestamp: bool = True,
        max_lines: int = 1000,
        **kwargs
    ):
        """
        初始化日誌面板
        
        Args:
            master: 父組件
            title: 面板標題
            show_timestamp: 是否顯示時間戳
            max_lines: 最大行數（超過自動清理舊記錄）
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
        self._show_timestamp = show_timestamp
        self._max_lines = max_lines
        self._line_count = 0
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # === 標題列 ===
        header = ctk.CTkFrame(self, fg_color="transparent", height=36)
        header.pack(fill="x", padx=12, pady=(8, 0))
        header.pack_propagate(False)
        
        ctk.CTkLabel(
            header,
            text=f"📋 {self._title}",
            font=Fonts.to_tuple(Fonts.LABEL),
            text_color=theme_manager.colors.text_primary
        ).pack(side="left")
        
        # 工具按鈕
        btn_frame = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame.pack(side="right")
        
        self.btn_clear = ctk.CTkButton(
            btn_frame,
            text="清除",
            width=50,
            height=24,
            font=Fonts.to_tuple(Fonts.CAPTION),
            fg_color="transparent",
            hover_color=theme_manager.colors.border,
            text_color=theme_manager.colors.text_muted,
            command=self.clear
        )
        self.btn_clear.pack(side="left", padx=(4, 0))
        
        self.btn_copy = ctk.CTkButton(
            btn_frame,
            text="複製",
            width=50,
            height=24,
            font=Fonts.to_tuple(Fonts.CAPTION),
            fg_color="transparent",
            hover_color=theme_manager.colors.border,
            text_color=theme_manager.colors.text_muted,
            command=self._copy_all
        )
        self.btn_copy.pack(side="left", padx=(4, 0))
        
        # === 日誌區域 ===
        log_frame = ctk.CTkFrame(self, fg_color="transparent")
        log_frame.pack(fill="both", expand=True, padx=8, pady=8)
        
        # 使用 Text widget（CustomTkinter 的 CTkTextbox 不支援標籤顏色）
        self.text = tk.Text(
            log_frame,
            wrap="word",
            font=Fonts.to_tuple(Fonts.CODE),
            bg=theme_manager.colors.background,
            fg=theme_manager.colors.text_primary,
            insertbackground=theme_manager.colors.text_primary,
            selectbackground=theme_manager.colors.primary,
            selectforeground="#ffffff",
            relief="flat",
            padx=8,
            pady=8,
            state="disabled",
            cursor="arrow"
        )
        self.text.pack(side="left", fill="both", expand=True)
        
        # 滾動條
        scrollbar = ctk.CTkScrollbar(log_frame, command=self.text.yview)
        scrollbar.pack(side="right", fill="y")
        self.text.configure(yscrollcommand=scrollbar.set)
        
        # 配置標籤顏色
        self._configure_tags()
    
    def _configure_tags(self):
        """配置文字標籤顏色"""
        for level, color in self.LEVEL_COLORS.items():
            self.text.tag_configure(level, foreground=color)
        
        # 時間戳標籤
        self.text.tag_configure("TIMESTAMP", foreground=theme_manager.colors.text_muted)
        
        # 預設標籤
        self.text.tag_configure("DEFAULT", foreground=theme_manager.colors.text_primary)
    
    def _detect_level(self, message: str) -> str:
        """自動檢測日誌級別"""
        message_lower = message.lower()
        
        for pattern, level in self.PATTERNS.items():
            if re.search(pattern, message_lower, re.IGNORECASE):
                return level
        
        return "DEFAULT"
    
    def _format_timestamp(self) -> str:
        """格式化時間戳"""
        return datetime.now().strftime("[%H:%M:%S]")
    
    def log(self, message: str, level: Optional[str] = None):
        """
        寫入日誌
        
        Args:
            message: 日誌訊息
            level: 日誌級別 (INFO, SUCCESS, WARNING, ERROR, DEBUG)
                   若為 None 則自動檢測
        """
        # 自動檢測級別
        if level is None:
            level = self._detect_level(message)
        
        # 確保級別有效
        if level not in self.LEVEL_COLORS and level != "DEFAULT":
            level = "DEFAULT"
        
        # 啟用編輯
        self.text.configure(state="normal")
        
        # 檢查是否需要清理舊記錄
        if self._line_count >= self._max_lines:
            self._trim_old_lines()
        
        # 添加時間戳
        if self._show_timestamp:
            timestamp = self._format_timestamp()
            self.text.insert("end", timestamp + " ", "TIMESTAMP")
        
        # 添加訊息
        self.text.insert("end", message + "\n", level)
        
        # 禁用編輯
        self.text.configure(state="disabled")
        
        # 自動滾動到底部
        self.text.see("end")
        
        self._line_count += 1
    
    def info(self, message: str):
        """資訊日誌"""
        self.log(message, "INFO")
    
    def success(self, message: str):
        """成功日誌"""
        self.log(message, "SUCCESS")
    
    def warning(self, message: str):
        """警告日誌"""
        self.log(message, "WARNING")
    
    def error(self, message: str):
        """錯誤日誌"""
        self.log(message, "ERROR")
    
    def debug(self, message: str):
        """除錯日誌"""
        self.log(message, "DEBUG")
    
    def _trim_old_lines(self):
        """清理舊記錄（保留最後 80%）"""
        keep_lines = int(self._max_lines * 0.8)
        delete_lines = self._line_count - keep_lines
        
        if delete_lines > 0:
            self.text.delete("1.0", f"{delete_lines + 1}.0")
            self._line_count = keep_lines
    
    def clear(self):
        """清除所有日誌"""
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")
        self._line_count = 0
    
    def _copy_all(self):
        """複製所有日誌到剪貼簿"""
        content = self.text.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(content)
    
    def get_content(self) -> str:
        """取得所有日誌內容"""
        return self.text.get("1.0", "end-1c")


class CollapsibleLogPanel(ctk.CTkFrame):
    """
    可收合的日誌面板
    整合標題列與日誌內容，點擊標題列可展開/收合
    """
    
    # 日誌級別對應顏色
    LEVEL_COLORS = {
        "INFO": "#3b82f6",      # 藍色
        "SUCCESS": "#22c55e",   # 綠色
        "WARNING": "#f59e0b",   # 黃色
        "ERROR": "#ef4444",     # 紅色
        "DEBUG": "#6b7280",     # 灰色
    }
    
    # 自動檢測模式
    PATTERNS = {
        r"^✓|成功|完成|done|success": "SUCCESS",
        r"^✗|失敗|錯誤|error|fail": "ERROR",
        r"警告|注意|warning": "WARNING",
        r"^\[DEBUG\]|debug": "DEBUG",
    }
    
    def __init__(
        self,
        master: Any,
        title: str = "系統日誌",
        default_expanded: bool = True,
        min_height: int = 150,
        show_timestamp: bool = True,
        max_lines: int = 1000,
        **kwargs
    ):
        """
        初始化可收合日誌面板
        
        Args:
            master: 父組件
            title: 面板標題
            default_expanded: 預設是否展開
            min_height: 展開時的最小高度
            show_timestamp: 是否顯示時間戳
            max_lines: 最大行數
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
        self._expanded = default_expanded
        self._min_height = min_height
        self._show_timestamp = show_timestamp
        self._max_lines = max_lines
        self._line_count = 0
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # === 標題列（可點擊收合） ===
        self.header = ctk.CTkFrame(self, fg_color="transparent", height=36)
        self.header.pack(fill="x", padx=8, pady=(6, 0))
        self.header.pack_propagate(False)
        
        # 左側：箭頭 + 標題（可點擊）
        left_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        left_frame.pack(side="left", fill="y")
        left_frame.bind("<Button-1>", lambda e: self.toggle())
        
        self.arrow_label = ctk.CTkLabel(
            left_frame,
            text="▼" if self._expanded else "▶",
            font=("Segoe UI", 10),
            text_color=theme_manager.colors.text_muted,
            width=16
        )
        self.arrow_label.pack(side="left")
        self.arrow_label.bind("<Button-1>", lambda e: self.toggle())
        
        self.title_label = ctk.CTkLabel(
            left_frame,
            text=f"📋 {self._title}",
            font=Fonts.to_tuple(Fonts.LABEL),
            text_color=theme_manager.colors.text_primary
        )
        self.title_label.pack(side="left", padx=(4, 0))
        self.title_label.bind("<Button-1>", lambda e: self.toggle())
        
        # 右側：工具按鈕
        btn_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        btn_frame.pack(side="right")
        
        self.btn_clear = ctk.CTkButton(
            btn_frame,
            text="清除",
            width=50,
            height=24,
            font=Fonts.to_tuple(Fonts.CAPTION),
            fg_color="transparent",
            hover_color=theme_manager.colors.border,
            text_color=theme_manager.colors.text_muted,
            command=self.clear
        )
        self.btn_clear.pack(side="left", padx=(4, 0))
        
        self.btn_copy = ctk.CTkButton(
            btn_frame,
            text="複製",
            width=50,
            height=24,
            font=Fonts.to_tuple(Fonts.CAPTION),
            fg_color="transparent",
            hover_color=theme_manager.colors.border,
            text_color=theme_manager.colors.text_muted,
            command=self._copy_all
        )
        self.btn_copy.pack(side="left", padx=(4, 0))
        
        # === 日誌內容區域 ===
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        if self._expanded:
            self.content_frame.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        
        # 使用 Text widget
        self.text = tk.Text(
            self.content_frame,
            wrap="word",
            font=Fonts.to_tuple(Fonts.CODE),
            bg=theme_manager.colors.background,
            fg=theme_manager.colors.text_primary,
            insertbackground=theme_manager.colors.text_primary,
            selectbackground=theme_manager.colors.primary,
            selectforeground="#ffffff",
            relief="flat",
            padx=8,
            pady=8,
            state="disabled",
            cursor="arrow",
            height=8  # 預設高度
        )
        self.text.pack(fill="both", expand=True)
        
        # 滾動條
        scrollbar = ctk.CTkScrollbar(self.content_frame, command=self.text.yview)
        scrollbar.pack(side="right", fill="y")
        self.text.configure(yscrollcommand=scrollbar.set)
        
        # 重新排列讓 text 在左邊
        self.text.pack_forget()
        scrollbar.pack_forget()
        self.text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 配置標籤顏色
        self._configure_tags()
    
    def _configure_tags(self):
        """配置文字標籤顏色"""
        for level, color in self.LEVEL_COLORS.items():
            self.text.tag_configure(level, foreground=color)
        self.text.tag_configure("TIMESTAMP", foreground=theme_manager.colors.text_muted)
        self.text.tag_configure("DEFAULT", foreground=theme_manager.colors.text_primary)
    
    def toggle(self):
        """切換展開/收合狀態"""
        self._expanded = not self._expanded
        
        if self._expanded:
            self.content_frame.pack(fill="both", expand=True, padx=8, pady=(4, 8))
            self.arrow_label.configure(text="▼")
        else:
            self.content_frame.pack_forget()
            self.arrow_label.configure(text="▶")
    
    def expand(self):
        """展開"""
        if not self._expanded:
            self.toggle()
    
    def collapse(self):
        """收合"""
        if self._expanded:
            self.toggle()
    
    def _detect_level(self, message: str) -> str:
        """自動檢測日誌級別"""
        import re
        message_lower = message.lower()
        
        for pattern, level in self.PATTERNS.items():
            if re.search(pattern, message_lower, re.IGNORECASE):
                return level
        
        return "DEFAULT"
    
    def _format_timestamp(self) -> str:
        """格式化時間戳"""
        return datetime.now().strftime("[%H:%M:%S]")
    
    def log(self, message: str, level: Optional[str] = None):
        """寫入日誌"""
        if level is None:
            level = self._detect_level(message)
        
        if level not in self.LEVEL_COLORS and level != "DEFAULT":
            level = "DEFAULT"
        
        self.text.configure(state="normal")
        
        if self._line_count >= self._max_lines:
            self._trim_old_lines()
        
        if self._show_timestamp:
            timestamp = self._format_timestamp()
            self.text.insert("end", timestamp + " ", "TIMESTAMP")
        
        self.text.insert("end", message + "\n", level)
        self.text.configure(state="disabled")
        self.text.see("end")
        
        self._line_count += 1
    
    def _trim_old_lines(self):
        """清理舊記錄"""
        keep_lines = int(self._max_lines * 0.8)
        delete_lines = self._line_count - keep_lines
        
        if delete_lines > 0:
            self.text.delete("1.0", f"{delete_lines + 1}.0")
            self._line_count = keep_lines
    
    def info(self, message: str):
        self.log(message, "INFO")
    
    def success(self, message: str):
        self.log(message, "SUCCESS")
    
    def warning(self, message: str):
        self.log(message, "WARNING")
        self.expand()
    
    def error(self, message: str):
        self.log(message, "ERROR")
        self.expand()
    
    def debug(self, message: str):
        self.log(message, "DEBUG")
    
    def clear(self):
        """清除所有日誌"""
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")
        self._line_count = 0
    
    def _copy_all(self):
        """複製所有日誌到剪貼簿"""
        content = self.text.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(content)
    
    def get_content(self) -> str:
        """取得所有日誌內容"""
        return self.text.get("1.0", "end-1c")
