# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主題管理模組
集中管理顏色、字體、樣式等視覺元素
"""

import customtkinter as ctk
from typing import Dict, Tuple, Optional
from dataclasses import dataclass


@dataclass
class ColorScheme:
    """顏色方案定義"""
    # 主要顏色
    primary: str
    primary_hover: str
    primary_disabled: str
    
    # 次要顏色
    secondary: str
    secondary_hover: str
    
    # 狀態顏色
    success: str
    success_hover: str
    warning: str
    warning_hover: str
    danger: str
    danger_hover: str
    info: str
    info_hover: str
    
    # 背景顏色
    bg_primary: str
    bg_secondary: str
    bg_tertiary: str
    bg_card: str
    bg_hover: str
    
    # 文字顏色
    text_primary: str
    text_secondary: str
    text_muted: str
    text_inverse: str
    
    # 邊框顏色
    border: str
    border_hover: str
    border_focus: str
    
    # 陰影
    shadow: str
    
    # 別名屬性（向後兼容）
    @property
    def background(self) -> str:
        return self.bg_primary
    
    @property
    def card_bg(self) -> str:
        return self.bg_card


class Colors:
    """預定義顏色常數"""
    
    # === 深色主題 (Dark Mode) ===
    DARK = ColorScheme(
        # 主要顏色 - 藍色系
        primary="#3B82F6",
        primary_hover="#2563EB",
        primary_disabled="#1E40AF",
        
        # 次要顏色 - 灰色系
        secondary="#64748B",
        secondary_hover="#475569",
        
        # 狀態顏色
        success="#22C55E",
        success_hover="#16A34A",
        warning="#F59E0B",
        warning_hover="#D97706",
        danger="#EF4444",
        danger_hover="#DC2626",
        info="#06B6D4",
        info_hover="#0891B2",
        
        # 背景顏色
        bg_primary="#0F172A",      # 最深背景
        bg_secondary="#1E293B",    # 次深背景
        bg_tertiary="#334155",     # 第三層背景
        bg_card="#1E293B",         # 卡片背景
        bg_hover="#374151",        # 懸停背景
        
        # 文字顏色
        text_primary="#F8FAFC",
        text_secondary="#CBD5E1",
        text_muted="#64748B",
        text_inverse="#0F172A",
        
        # 邊框顏色
        border="#334155",
        border_hover="#475569",
        border_focus="#3B82F6",
        
        # 陰影
        shadow="#00000040",
    )
    
    # === 淺色主題 (Light Mode) ===
    LIGHT = ColorScheme(
        # 主要顏色 - 藍色系
        primary="#2563EB",
        primary_hover="#1D4ED8",
        primary_disabled="#93C5FD",
        
        # 次要顏色 - 灰色系
        secondary="#6B7280",
        secondary_hover="#4B5563",
        
        # 狀態顏色
        success="#16A34A",
        success_hover="#15803D",
        warning="#D97706",
        warning_hover="#B45309",
        danger="#DC2626",
        danger_hover="#B91C1C",
        info="#0891B2",
        info_hover="#0E7490",
        
        # 背景顏色
        bg_primary="#F8FAFC",      # 最淺背景
        bg_secondary="#F1F5F9",    # 次淺背景
        bg_tertiary="#E2E8F0",     # 第三層背景
        bg_card="#FFFFFF",         # 卡片背景
        bg_hover="#F1F5F9",        # 懸停背景
        
        # 文字顏色
        text_primary="#0F172A",
        text_secondary="#475569",
        text_muted="#94A3B8",
        text_inverse="#F8FAFC",
        
        # 邊框顏色
        border="#E2E8F0",
        border_hover="#CBD5E1",
        border_focus="#2563EB",
        
        # 陰影
        shadow="#00000015",
    )


@dataclass
class FontConfig:
    """字體配置"""
    family: str
    size: int
    weight: str = "normal"


class Fonts:
    """字體預設配置"""
    
    # 標題字體
    TITLE_LARGE = FontConfig("Microsoft JhengHei UI", 24, "bold")
    TITLE = FontConfig("Microsoft JhengHei UI", 18, "bold")
    TITLE_SMALL = FontConfig("Microsoft JhengHei UI", 14, "bold")
    
    # 內文字體
    BODY_LARGE = FontConfig("Microsoft JhengHei UI", 14, "normal")
    BODY = FontConfig("Microsoft JhengHei UI", 12, "normal")
    BODY_SMALL = FontConfig("Microsoft JhengHei UI", 11, "normal")
    
    # 輔助字體
    CAPTION = FontConfig("Microsoft JhengHei UI", 10, "normal")
    CAPTION_SMALL = FontConfig("Microsoft JhengHei UI", 9, "normal")
    
    # 按鈕字體
    BUTTON = FontConfig("Microsoft JhengHei UI", 12, "bold")
    BUTTON_SMALL = FontConfig("Microsoft JhengHei UI", 11, "normal")
    
    # 程式碼/日誌字體
    CODE = FontConfig("Consolas", 11, "normal")
    CODE_SMALL = FontConfig("Consolas", 10, "normal")
    
    # 標籤字體
    LABEL = FontConfig("Microsoft JhengHei UI", 12, "normal")
    LABEL_SMALL = FontConfig("Microsoft JhengHei UI", 10, "normal")
    
    @staticmethod
    def to_tuple(font_config: FontConfig) -> Tuple[str, int, str]:
        """轉換為 tkinter 字體元組"""
        return (font_config.family, font_config.size, font_config.weight)


class Spacing:
    """間距預設值"""
    
    # 基礎間距單位
    UNIT = 4
    
    # 常用間距
    XS = 4
    SM = 8
    MD = 12
    LG = 16
    XL = 24
    XXL = 32
    
    # Padding
    PADDING_CARD = (16, 16)
    PADDING_BUTTON = (16, 8)
    PADDING_INPUT = (12, 8)
    PADDING_PAGE = (24, 20)
    
    # Border radius
    RADIUS_SM = 4
    RADIUS_MD = 8
    RADIUS_LG = 12
    RADIUS_XL = 16
    RADIUS_FULL = 9999


class ThemeManager:
    """
    主題管理器
    負責管理應用程式的整體主題設定
    """
    
    _instance: Optional['ThemeManager'] = None
    _current_theme: str = "dark"
    _colors: ColorScheme = Colors.DARK
    
    def __new__(cls) -> 'ThemeManager':
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not hasattr(self, '_initialized'):
            self._initialized = True
            self._theme_change_callbacks = []
            self._setup_customtkinter()
    
    def _setup_customtkinter(self):
        """設定 CustomTkinter 預設值"""
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
    
    @property
    def colors(self) -> ColorScheme:
        """取得當前顏色方案"""
        return self._colors
    
    @property
    def is_dark(self) -> bool:
        """是否為深色主題"""
        return self._current_theme == "dark"
    
    def set_theme(self, theme: str):
        """
        設定主題
        
        Args:
            theme: "dark" 或 "light"
        """
        if theme not in ("dark", "light"):
            raise ValueError(f"Invalid theme: {theme}")
        
        self._current_theme = theme
        self._colors = Colors.DARK if theme == "dark" else Colors.LIGHT
        
        # 更新 CustomTkinter
        ctk.set_appearance_mode(theme)
        
        # 通知所有監聽器
        for callback in self._theme_change_callbacks:
            try:
                callback(theme)
            except Exception:
                pass
    
    def toggle_theme(self):
        """切換深淺主題"""
        new_theme = "light" if self.is_dark else "dark"
        self.set_theme(new_theme)
    
    def add_theme_change_callback(self, callback):
        """添加主題變更回調"""
        if callback not in self._theme_change_callbacks:
            self._theme_change_callbacks.append(callback)
    
    def remove_theme_change_callback(self, callback):
        """移除主題變更回調"""
        if callback in self._theme_change_callbacks:
            self._theme_change_callbacks.remove(callback)
    
    # === 便捷方法 ===
    
    def get_button_colors(self, variant: str = "primary") -> Dict:
        """取得按鈕顏色配置"""
        c = self._colors
        
        variants = {
            "primary": {
                "fg_color": c.primary,
                "hover_color": c.primary_hover,
                "text_color": c.text_inverse,
            },
            "secondary": {
                "fg_color": c.secondary,
                "hover_color": c.secondary_hover,
                "text_color": c.text_inverse,
            },
            "success": {
                "fg_color": c.success,
                "hover_color": c.success_hover,
                "text_color": c.text_inverse,
            },
            "warning": {
                "fg_color": c.warning,
                "hover_color": c.warning_hover,
                "text_color": c.text_inverse,
            },
            "danger": {
                "fg_color": c.danger,
                "hover_color": c.danger_hover,
                "text_color": c.text_inverse,
            },
            "outline": {
                "fg_color": "transparent",
                "hover_color": c.bg_hover,
                "text_color": c.text_primary,
                "border_color": c.border,
            },
            "ghost": {
                "fg_color": "transparent",
                "hover_color": c.bg_hover,
                "text_color": c.text_primary,
            },
        }
        
        return variants.get(variant, variants["primary"])
    
    def get_card_style(self) -> Dict:
        """取得卡片樣式"""
        c = self._colors
        return {
            "fg_color": c.bg_card,
            "border_color": c.border,
            "corner_radius": Spacing.RADIUS_LG,
        }
    
    def get_input_style(self) -> Dict:
        """取得輸入框樣式"""
        c = self._colors
        return {
            "fg_color": c.bg_secondary,
            "border_color": c.border,
            "text_color": c.text_primary,
            "placeholder_text_color": c.text_muted,
            "corner_radius": Spacing.RADIUS_MD,
        }


# 全域主題管理器實例
theme_manager = ThemeManager()
