# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
現代化 UI 組件模組
提供可重用的現代化 UI 組件
"""

import customtkinter as ctk
import tkinter as tk
from typing import Optional, Callable, List, Union, Any
from ui.theme import ThemeManager, Fonts, Spacing, theme_manager


class ModernButton(ctk.CTkButton):
    """
    現代化按鈕組件
    
    支援多種樣式變體：primary, secondary, success, warning, danger, outline, ghost
    """
    
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
        
        # 添加 Tooltip
        if tooltip:
            ModernTooltip(self, tooltip)


class ModernEntry(ctk.CTkEntry):
    """
    現代化輸入框組件
    """
    
    def __init__(
        self,
        master: Any,
        placeholder: str = "",
        label: Optional[str] = None,
        width: int = 200,
        show_clear_button: bool = False,
        on_change: Optional[Callable[[str], None]] = None,
        **kwargs
    ):
        """
        初始化輸入框
        
        Args:
            master: 父組件
            placeholder: 佔位文字
            label: 標籤文字
            width: 寬度
            show_clear_button: 是否顯示清除按鈕
            on_change: 內容變更回調
        """
        style = theme_manager.get_input_style()
        
        super().__init__(
            master,
            width=width,
            height=36,
            placeholder_text=placeholder,
            font=Fonts.to_tuple(Fonts.BODY),
            **style,
            **kwargs
        )
        
        self.on_change = on_change
        
        # 綁定變更事件
        if on_change:
            self.bind("<KeyRelease>", self._handle_change)
    
    def _handle_change(self, event=None):
        """處理內容變更"""
        if self.on_change:
            self.on_change(self.get())


class ModernCard(ctk.CTkFrame):
    """
    現代化卡片組件
    用於組織和分組相關內容
    """
    
    def __init__(
        self,
        master: Any,
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
        padding: int = 16,
        **kwargs
    ):
        """
        初始化卡片
        
        Args:
            master: 父組件
            title: 標題
            subtitle: 副標題
            padding: 內邊距
        """
        style = theme_manager.get_card_style()
        
        super().__init__(
            master,
            border_width=1,
            **style,
            **kwargs
        )
        
        self._padding = padding
        
        # 內容容器
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True, padx=padding, pady=padding)
        
        # 標題區域
        if title or subtitle:
            self._create_header(title, subtitle)
    
    def _create_header(self, title: Optional[str], subtitle: Optional[str]):
        """建立標題區域"""
        header_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 12))
        
        if title:
            title_label = ctk.CTkLabel(
                header_frame,
                text=title,
                font=Fonts.to_tuple(Fonts.TITLE_SMALL),
                text_color=theme_manager.colors.text_primary
            )
            title_label.pack(anchor="w")
        
        if subtitle:
            subtitle_label = ctk.CTkLabel(
                header_frame,
                text=subtitle,
                font=Fonts.to_tuple(Fonts.BODY_SMALL),
                text_color=theme_manager.colors.text_secondary
            )
            subtitle_label.pack(anchor="w")
    
    def get_content_frame(self) -> ctk.CTkFrame:
        """取得內容框架（用於添加子組件）"""
        return self.content_frame


class ModernTooltip:
    """
    現代化 Tooltip 組件
    """
    
    def __init__(
        self,
        widget: Any,
        text: Union[str, List[str]],
        delay: int = 500,
        duration: int = 4000
    ):
        """
        初始化 Tooltip
        
        Args:
            widget: 目標組件
            text: 提示文字（字串或行列表）
            delay: 顯示延遲（毫秒）
            duration: 顯示持續時間（毫秒）
        """
        self.widget = widget
        self.text = text if isinstance(text, str) else "\n".join(text)
        self.delay = delay
        self.duration = duration
        self.tooltip_window: Optional[ctk.CTkToplevel] = None
        self._show_timer: Optional[str] = None
        self._hide_timer: Optional[str] = None
        
        # 綁定事件
        widget.bind("<Enter>", self._on_enter)
        widget.bind("<Leave>", self._on_leave)
    
    def _on_enter(self, event):
        """滑鼠進入"""
        self._cancel_timers()
        self._show_timer = self.widget.after(self.delay, self._show)
    
    def _on_leave(self, event):
        """滑鼠離開"""
        self._cancel_timers()
        self._hide()
    
    def _cancel_timers(self):
        """取消所有計時器"""
        if self._show_timer:
            self.widget.after_cancel(self._show_timer)
            self._show_timer = None
        if self._hide_timer:
            self.widget.after_cancel(self._hide_timer)
            self._hide_timer = None
    
    def _show(self):
        """顯示 Tooltip"""
        if self.tooltip_window:
            return
        
        # 取得位置
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        
        # 建立視窗
        self.tooltip_window = ctk.CTkToplevel()
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        
        # 設定外觀
        self.tooltip_window.configure(fg_color=theme_manager.colors.bg_tertiary)
        
        # 內容標籤
        label = ctk.CTkLabel(
            self.tooltip_window,
            text=self.text,
            font=Fonts.to_tuple(Fonts.BODY_SMALL),
            text_color=theme_manager.colors.text_primary,
            justify="left",
            padx=12,
            pady=8
        )
        label.pack()
        
        # 自動隱藏計時器
        self._hide_timer = self.widget.after(self.duration, self._hide)
    
    def _hide(self):
        """隱藏 Tooltip"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
    
    def update_text(self, text: Union[str, List[str]]):
        """更新提示文字"""
        self.text = text if isinstance(text, str) else "\n".join(text)


class StatusBadge(ctk.CTkFrame):
    """
    狀態標籤組件
    用於顯示狀態指示器
    """
    
    STATUS_COLORS = {
        "success": ("#22C55E", "#16A34A", "#D1FAE5"),  # color, hover, bg
        "warning": ("#F59E0B", "#D97706", "#FEF3C7"),
        "danger": ("#EF4444", "#DC2626", "#FEE2E2"),
        "info": ("#06B6D4", "#0891B2", "#CFFAFE"),
        "default": ("#64748B", "#475569", "#E2E8F0"),
    }
    
    def __init__(
        self,
        master: Any,
        text: str,
        status: str = "default",
        show_dot: bool = True,
        **kwargs
    ):
        """
        初始化狀態標籤
        
        Args:
            master: 父組件
            text: 狀態文字
            status: 狀態類型 (success/warning/danger/info/default)
            show_dot: 是否顯示狀態點
        """
        color_light, color_dark, bg_color = self.STATUS_COLORS.get(status, self.STATUS_COLORS["default"])
        
        # 深色模式使用較暗的背景
        if ctk.get_appearance_mode() == "Dark":
            bg_color = color_dark
        
        super().__init__(
            master,
            fg_color=bg_color,
            corner_radius=Spacing.RADIUS_SM,
            **kwargs
        )
        
        inner_frame = ctk.CTkFrame(self, fg_color="transparent")
        inner_frame.pack(padx=8, pady=4)
        
        # 狀態點
        if show_dot:
            dot_frame = ctk.CTkFrame(
                inner_frame,
                width=8,
                height=8,
                corner_radius=4,
                fg_color=color_light
            )
            dot_frame.pack(side="left", padx=(0, 6))
        
        # 狀態文字
        text_color = "#FFFFFF" if ctk.get_appearance_mode() == "Dark" else color_dark
        label = ctk.CTkLabel(
            inner_frame,
            text=text,
            font=Fonts.to_tuple(Fonts.LABEL_SMALL),
            text_color=text_color
        )
        label.pack(side="left")
        
        self._status = status
        self._label = label
        self._color = color_light
    
    def set_status(self, text: str, status: str):
        """更新狀態"""
        color_light, color_dark, bg_color = self.STATUS_COLORS.get(status, self.STATUS_COLORS["default"])
        if ctk.get_appearance_mode() == "Dark":
            bg_color = color_dark
            text_color = "#FFFFFF"
        else:
            text_color = color_dark
        self.configure(fg_color=bg_color)
        self._label.configure(text=text, text_color=text_color)


class ModernProgressBar(ctk.CTkProgressBar):
    """
    現代化進度條組件
    """
    
    def __init__(
        self,
        master: Any,
        width: int = 200,
        height: int = 8,
        mode: str = "determinate",
        **kwargs
    ):
        super().__init__(
            master,
            width=width,
            height=height,
            mode=mode,
            progress_color=theme_manager.colors.primary,
            fg_color=theme_manager.colors.bg_tertiary,
            corner_radius=height // 2,
            **kwargs
        )


class ModernSwitch(ctk.CTkSwitch):
    """
    現代化開關組件
    """
    
    def __init__(
        self,
        master: Any,
        text: str = "",
        command: Optional[Callable] = None,
        variable: Optional[ctk.Variable] = None,
        **kwargs
    ):
        super().__init__(
            master,
            text=text,
            command=command,
            variable=variable,
            font=Fonts.to_tuple(Fonts.BODY),
            text_color=theme_manager.colors.text_primary,
            progress_color=theme_manager.colors.primary,
            button_color=theme_manager.colors.text_inverse,
            button_hover_color=theme_manager.colors.bg_hover,
            fg_color=theme_manager.colors.bg_tertiary,
            **kwargs
        )


class FormField(ctk.CTkFrame):
    """
    表單欄位組件
    包含標籤和輸入框的組合
    """
    
    def __init__(
        self,
        master: Any,
        label: str,
        label_width: int = 100,
        input_width: int = 300,
        placeholder: str = "",
        variable: Optional[ctk.StringVar] = None,
        **kwargs
    ):
        super().__init__(master, fg_color="transparent", **kwargs)
        
        # 標籤
        self.label = ctk.CTkLabel(
            self,
            text=label,
            font=Fonts.to_tuple(Fonts.LABEL),
            text_color=theme_manager.colors.text_primary,
            width=label_width,
            anchor="w"
        )
        self.label.pack(side="left")
        
        # 輸入框
        self.entry = ModernEntry(
            self,
            placeholder=placeholder,
            width=input_width,
            textvariable=variable
        )
        self.entry.pack(side="left", padx=(8, 0), fill="x", expand=True)
    
    def get(self) -> str:
        """取得輸入值"""
        return self.entry.get()
    
    def set(self, value: str):
        """設定輸入值"""
        self.entry.delete(0, "end")
        self.entry.insert(0, value)


class SectionTitle(ctk.CTkFrame):
    """
    區段標題組件
    """
    
    def __init__(
        self,
        master: Any,
        title: str,
        subtitle: Optional[str] = None,
        action_button: Optional[dict] = None,
        **kwargs
    ):
        super().__init__(master, fg_color="transparent", **kwargs)
        
        # 左側：標題
        left_frame = ctk.CTkFrame(self, fg_color="transparent")
        left_frame.pack(side="left", fill="x", expand=True)
        
        title_label = ctk.CTkLabel(
            left_frame,
            text=title,
            font=Fonts.to_tuple(Fonts.TITLE_SMALL),
            text_color=theme_manager.colors.text_primary
        )
        title_label.pack(anchor="w")
        
        if subtitle:
            subtitle_label = ctk.CTkLabel(
                left_frame,
                text=subtitle,
                font=Fonts.to_tuple(Fonts.BODY_SMALL),
                text_color=theme_manager.colors.text_muted
            )
            subtitle_label.pack(anchor="w")
        
        # 右側：操作按鈕
        if action_button:
            btn = ModernButton(
                self,
                text=action_button.get("text", ""),
                command=action_button.get("command"),
                variant=action_button.get("variant", "ghost"),
                size="sm"
            )
            btn.pack(side="right")


class EmptyState(ctk.CTkFrame):
    """
    空狀態組件
    當列表為空時顯示
    """
    
    def __init__(
        self,
        master: Any,
        icon: str = "📭",
        title: str = "沒有資料",
        description: str = "",
        action_button: Optional[dict] = None,
        **kwargs
    ):
        super().__init__(master, fg_color="transparent", **kwargs)
        
        # 圖示
        icon_label = ctk.CTkLabel(
            self,
            text=icon,
            font=("Segoe UI Emoji", 48)
        )
        icon_label.pack(pady=(20, 10))
        
        # 標題
        title_label = ctk.CTkLabel(
            self,
            text=title,
            font=Fonts.to_tuple(Fonts.TITLE_SMALL),
            text_color=theme_manager.colors.text_primary
        )
        title_label.pack()
        
        # 描述
        if description:
            desc_label = ctk.CTkLabel(
                self,
                text=description,
                font=Fonts.to_tuple(Fonts.BODY),
                text_color=theme_manager.colors.text_muted
            )
            desc_label.pack(pady=(4, 0))
        
        # 操作按鈕
        if action_button:
            btn = ModernButton(
                self,
                text=action_button.get("text", ""),
                command=action_button.get("command"),
                variant=action_button.get("variant", "primary"),
                icon=action_button.get("icon")
            )
            btn.pack(pady=(16, 0))
