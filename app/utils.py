# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用工具模組
- UI 輔助函數
- Tooltip 工具
- 其他通用函數
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Optional, Callable


class Tooltip:
    """
    通用 Tooltip 工具類別
    用於為任何 tkinter widget 添加滑鼠懸停提示
    """
    
    def __init__(self, widget: tk.Widget, lines: List[str], delay: int = 4000):
        """
        初始化 Tooltip
        
        Args:
            widget: 要添加 tooltip 的 widget
            lines: 要顯示的文字行列表
            delay: 自動隱藏延遲（毫秒）
        """
        self.widget = widget
        self.lines = lines
        self.delay = delay
        self.tooltip_window: Optional[tk.Toplevel] = None
        
        # 綁定事件
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)
    
    def _show(self, event):
        """顯示 tooltip"""
        # 如果已存在，先關閉
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
        
        # 建立 tooltip 視窗
        self.tooltip_window = tk.Toplevel()
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
        
        # 建立框架
        frame = tk.Frame(self.tooltip_window, bg="lightyellow", relief="solid", bd=1)
        frame.pack()
        
        # 添加文字行
        for line in self.lines:
            label = tk.Label(
                frame, 
                text=line, 
                bg="lightyellow",
                font=("Arial", 9), 
                anchor="w", 
                justify="left"
            )
            label.pack(anchor="w", padx=8, pady=1)
        
        # 設定自動隱藏
        self.tooltip_window.after(self.delay, self._auto_hide)
    
    def _hide(self, event=None):
        """隱藏 tooltip"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
    
    def _auto_hide(self):
        """自動隱藏（計時器觸發）"""
        self._hide()
    
    def update_lines(self, lines: List[str]):
        """更新顯示的文字"""
        self.lines = lines


def create_smart_fix_tooltip(widget: tk.Widget) -> Tooltip:
    """
    為一鍵修復按鈕創建預設的 tooltip
    
    Args:
        widget: 按鈕 widget
        
    Returns:
        Tooltip 實例
    """
    lines = [
        "🔧 智能一鍵修復",
        "自動診斷並修復所有 WIM 掛載問題",
        "",
        "包含功能：",
        "• 狀態檢查與診斷",
        "• 清理掛載衝突",
        "• 修復損壞掛載",
        "• 系統級清理"
    ]
    return Tooltip(widget, lines, delay=4000)


def center_window(window: tk.Toplevel, width: int, height: int, parent: Optional[tk.Tk] = None):
    """
    將視窗置中顯示
    
    Args:
        window: 要置中的視窗
        width: 視窗寬度
        height: 視窗高度
        parent: 父視窗（如果提供，則相對於父視窗置中）
    """
    window.update_idletasks()
    
    if parent:
        x = parent.winfo_x() + (parent.winfo_width() - width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - height) // 2
    else:
        x = (window.winfo_screenwidth() - width) // 2
        y = (window.winfo_screenheight() - height) // 2
    
    window.geometry(f"{width}x{height}+{x}+{y}")


def create_labeled_entry(
    parent: tk.Widget,
    label_text: str,
    label_width: int = 12,
    entry_width: int = 40,
    variable: Optional[tk.StringVar] = None
) -> tuple:
    """
    創建帶標籤的輸入框
    
    Args:
        parent: 父 widget
        label_text: 標籤文字
        label_width: 標籤寬度
        entry_width: 輸入框寬度
        variable: StringVar 變數
        
    Returns:
        (frame, entry, variable) 元組
    """
    frame = ttk.Frame(parent)
    
    ttk.Label(frame, text=label_text, width=label_width).pack(side=tk.LEFT)
    
    if variable is None:
        variable = tk.StringVar()
    
    entry = ttk.Entry(frame, textvariable=variable, width=entry_width)
    entry.pack(side=tk.LEFT, padx=(8, 6), fill=tk.X, expand=True)
    
    return frame, entry, variable


def create_button_group(parent: tk.Widget, buttons: List[dict]) -> ttk.Frame:
    """
    創建按鈕組
    
    Args:
        parent: 父 widget
        buttons: 按鈕配置列表，每個元素為 dict:
                 {"text": str, "command": callable, "width": int (optional)}
                 
    Returns:
        包含所有按鈕的 Frame
    """
    frame = ttk.Frame(parent)
    
    for i, btn_config in enumerate(buttons):
        text = btn_config.get("text", "")
        command = btn_config.get("command")
        width = btn_config.get("width")
        
        btn = ttk.Button(frame, text=text, command=command)
        if width:
            btn.configure(width=width)
        
        padx = (8, 0) if i > 0 else 0
        btn.pack(side=tk.LEFT, padx=padx)
    
    return frame


def format_file_size(size_bytes: int) -> str:
    """
    格式化檔案大小
    
    Args:
        size_bytes: 位元組數
        
    Returns:
        格式化的字串 (如 "1.5 GB")
    """
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} PB"


def safe_destroy(widget: Optional[tk.Widget]):
    """
    安全地銷毀 widget
    
    Args:
        widget: 要銷毀的 widget
    """
    if widget:
        try:
            widget.destroy()
        except tk.TclError:
            pass  # Widget 已經被銷毀
