# -*- coding: utf-8 -*-
"""
回饋組件模組

提供 StatusBadge, ModernProgressBar, ModernTooltip, ModernSwitch 組件
"""

import customtkinter as ctk
from typing import Optional, Callable, List, Union, Any
from ui.theme import Fonts, Spacing, theme_manager


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
        "default": ("#94a3b8", "#64748b", "#f1f5f9"),
    }
    
    def __init__(
        self,
        master: Any,
        text: str,
        status: str = "default",
        show_dot: bool = True,
        animated: bool = True,
        **kwargs
    ):
        """
        初始化狀態標籤
        
        Args:
            master: 父組件
            text: 狀態文字
            status: 狀態類型 (success/warning/danger/info/default)
            show_dot: 是否顯示狀態點
            animated: 是否啟用動畫效果
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
        inner_frame.pack(padx=10, pady=6)
        
        # 狀態點
        self._dot_frame = None
        if show_dot:
            self._dot_frame = ctk.CTkFrame(
                inner_frame,
                width=10,
                height=10,
                corner_radius=5,
                fg_color=color_light
            )
            self._dot_frame.pack(side="left", padx=(0, 8))
        
        # 狀態文字
        text_color = "#FFFFFF" if ctk.get_appearance_mode() == "Dark" else color_dark
        self._label = ctk.CTkLabel(
            inner_frame,
            text=text,
            font=Fonts.to_tuple(Fonts.LABEL),
            text_color=text_color
        )
        self._label.pack(side="left")
        
        self._status = status
        self._color = color_light
        self._animated = animated
        self._pulse_job = None
        self._pulse_state = True
        
        # 如果是 success 狀態且啟用動畫，開始脈動效果
        if animated and status == "success" and self._dot_frame:
            self._start_pulse()
    
    def _start_pulse(self):
        """開始脈動動畫"""
        if self._pulse_job:
            self.after_cancel(self._pulse_job)
        self._do_pulse()
    
    def _stop_pulse(self):
        """停止脈動動畫"""
        if self._pulse_job:
            self.after_cancel(self._pulse_job)
            self._pulse_job = None
    
    def _do_pulse(self):
        """執行脈動動畫"""
        if not self._dot_frame or not self.winfo_exists():
            return
        
        try:
            # 切換透明度
            if self._pulse_state:
                self._dot_frame.configure(fg_color=self._color)
            else:
                # 淡化效果
                color_light, _, bg_color = self.STATUS_COLORS.get(self._status, self.STATUS_COLORS["default"])
                self._dot_frame.configure(fg_color=bg_color)
            
            self._pulse_state = not self._pulse_state
            self._pulse_job = self.after(800, self._do_pulse)
        except:
            pass
    
    def set_status(self, text: str, status: str):
        """更新狀態"""
        self._stop_pulse()
        
        color_light, color_dark, bg_color = self.STATUS_COLORS.get(status, self.STATUS_COLORS["default"])
        if ctk.get_appearance_mode() == "Dark":
            bg_color = color_dark
            text_color = "#FFFFFF"
        else:
            text_color = color_dark
        self.configure(fg_color=bg_color)
        self._label.configure(text=text, text_color=text_color)
        
        self._status = status
        self._color = color_light
        
        # 更新狀態點顏色
        if self._dot_frame:
            self._dot_frame.configure(fg_color=color_light)
        
        # 如果是 success 狀態且啟用動畫，開始脈動效果
        if self._animated and status == "success" and self._dot_frame:
            self._start_pulse()


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
