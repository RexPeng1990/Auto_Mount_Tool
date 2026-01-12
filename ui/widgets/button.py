# -*- coding: utf-8 -*-
"""
按鈕組件模組

提供 ModernButton 和 IconButton 組件
"""

import customtkinter as ctk
from typing import Optional, Callable, Any
from ui.theme import Fonts, Spacing, theme_manager


class ModernButton(ctk.CTkButton):
    """
    現代化按鈕組件
    
    支援多種樣式變體：primary, secondary, success, warning, danger, outline, ghost
    """
    
    # 禁用時的灰色背景
    DISABLED_BG_COLOR = "#d0d0d0"
    DISABLED_HOVER_COLOR = "#d0d0d0"
    DISABLED_TEXT_COLOR = "#ffffff"
    
    def __init__(
        self,
        master: Any,
        text: str = "",
        command: Optional[Callable] = None,
        variant: str = "primary",
        size: str = "md",
        icon: Optional[str] = None,
        width: Optional[int] = None,
        **kwargs
    ):
        """
        初始化按鈕
        
        Args:
            master: 父組件
            text: 按鈕文字
            command: 點擊回調
            variant: 樣式變體 (primary/secondary/success/warning/danger/outline/ghost)
            size: 尺寸 (sm/md/lg)
            icon: 圖示文字（如 emoji）
            width: 寬度
        """
        self.variant = variant
        self.size = size
        
        # 取得顏色配置
        colors = theme_manager.get_button_colors(variant)
        self._original_colors = colors.copy()  # 儲存原始顏色
        self._text_color = colors.get("text_color", "#ffffff")
        
        # 設定尺寸
        size_config = self._get_size_config(size)
        
        # 組合文字和圖示
        display_text = f"{icon} {text}" if icon else text
        
        # 預設參數
        default_kwargs = {
            "text": display_text,
            "command": command,
            "font": Fonts.to_tuple(Fonts.BUTTON if size != "sm" else Fonts.BUTTON_SMALL),
            "corner_radius": Spacing.RADIUS_MD,
            **colors,
            **size_config,
        }
        
        # 覆蓋寬度
        if width:
            default_kwargs["width"] = width
        
        # 合併用戶參數
        default_kwargs.update(kwargs)
        
        super().__init__(master, **default_kwargs)
    
    def configure(self, **kwargs):
        """覆寫 configure 以處理禁用狀態的樣式"""
        if "state" in kwargs:
            state = kwargs["state"]
            if state == "disabled":
                # 禁用時設置灰色背景，白色文字
                kwargs.setdefault("fg_color", self.DISABLED_BG_COLOR)
                kwargs.setdefault("hover_color", self.DISABLED_HOVER_COLOR)
                kwargs.setdefault("text_color", self.DISABLED_TEXT_COLOR)
            elif state == "normal":
                # 恢復原始顏色
                kwargs.setdefault("fg_color", self._original_colors.get("fg_color"))
                kwargs.setdefault("hover_color", self._original_colors.get("hover_color"))
                kwargs.setdefault("text_color", self._original_colors.get("text_color"))
        
        super().configure(**kwargs)
    
    def _get_size_config(self, size: str) -> dict:
        """取得尺寸配置"""
        sizes = {
            "sm": {"height": 28, "width": 80},
            "md": {"height": 36, "width": 100},
            "lg": {"height": 44, "width": 120},
        }
        return sizes.get(size, sizes["md"])
    
    def set_loading(self, loading: bool = True):
        """設定載入狀態"""
        if loading:
            self._original_text = self.cget("text")
            self.configure(text="⏳ 處理中...", state="disabled")
        else:
            if hasattr(self, '_original_text'):
                self.configure(text=self._original_text, state="normal")


class IconButton(ctk.CTkButton):
    """
    圖示按鈕（僅圖示，無文字）
    """
    
    def __init__(
        self,
        master: Any,
        icon: str,
        command: Optional[Callable] = None,
        size: int = 32,
        variant: str = "ghost",
        tooltip: Optional[str] = None,
        **kwargs
    ):
        colors = theme_manager.get_button_colors(variant)
        
        super().__init__(
            master,
            text=icon,
            command=command,
            width=size,
            height=size,
            font=("Segoe UI Emoji", size // 2),
            corner_radius=Spacing.RADIUS_MD,
            **colors,
            **kwargs
        )
        
        # 添加 Tooltip (延遲匯入避免循環依賴)
        if tooltip:
            from ui.widgets.feedback import ModernTooltip
            ModernTooltip(self, tooltip)
