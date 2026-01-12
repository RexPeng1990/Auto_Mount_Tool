# -*- coding: utf-8 -*-
"""
表單組件模組

提供 ModernEntry, FormField 組件
"""

import customtkinter as ctk
from typing import Optional, Callable, Any
from ui.theme import Fonts, theme_manager


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


class ModernComboBox(ctk.CTkFrame):
    """
    現代化下拉選單組件
    使用更美觀的下拉箭頭
    """
    
    def __init__(
        self,
        master: Any,
        values: list = None,
        variable: Optional[ctk.StringVar] = None,
        width: int = 140,
        height: int = 32,
        command: Optional[Callable] = None,
        state: str = "normal",
        **kwargs
    ):
        # 移除不應傳給 CTkFrame 的參數
        kwargs.pop('fg_color', None)
        kwargs.pop('border_color', None)
        kwargs.pop('border_width', None)
        kwargs.pop('button_color', None)
        kwargs.pop('dropdown_fg_color', None)
        kwargs.pop('dropdown_hover_color', None)
        kwargs.pop('font', None)
        kwargs.pop('dropdown_font', None)
        
        # 設定自身的固定寬度，防止被擠壓
        super().__init__(master, fg_color="transparent", width=width, height=height, **kwargs)
        
        self._values = values or []
        self._variable = variable or ctk.StringVar()
        self._command = command
        self._state = state
        self._width = width
        self._height = height
        
        # 防止自身被擠壓
        self.pack_propagate(False)
        self.grid_propagate(False)
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # 外框容器
        self.container = ctk.CTkFrame(
            self,
            fg_color="#ffffff",
            corner_radius=6,
            border_width=1,
            border_color="#cce5ff",
            width=self._width,
            height=self._height
        )
        self.container.pack(fill="both", expand=True)
        self.container.pack_propagate(False)
        
        # 內部容器（用於放置文字和箭頭）
        inner = ctk.CTkFrame(self.container, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=2, pady=2)
        
        # 顯示文字
        self.label = ctk.CTkLabel(
            inner,
            textvariable=self._variable,
            font=Fonts.to_tuple(Fonts.BODY_SMALL),
            text_color="#37474f",
            anchor="w"
        )
        self.label.pack(side="left", fill="x", expand=True, padx=(8, 0))
        
        # 下拉箭頭按鈕 - 使用更美觀的符號
        self.arrow_btn = ctk.CTkButton(
            inner,
            text="▼",  # 下拉箭頭
            font=("Microsoft JhengHei UI", 10),
            width=26,
            height=self._height - 6,
            corner_radius=4,
            fg_color="#e1f0ff",
            hover_color="#bbdefb",
            text_color="#5c6bc0",
            command=self._toggle_dropdown
        )
        self.arrow_btn.pack(side="right", padx=1)
        
        # 綁定點擊事件
        self.container.bind("<Button-1>", lambda e: self._toggle_dropdown())
        self.label.bind("<Button-1>", lambda e: self._toggle_dropdown())
        
        # 下拉選單（初始隱藏）
        self._dropdown = None
        self._click_binding = None
    
    def _toggle_dropdown(self):
        """切換下拉選單顯示/隱藏"""
        if self._state == "disabled":
            return
        
        # 如果已有下拉選單，關閉它
        if self._dropdown and self._dropdown.winfo_exists():
            self._close_dropdown()
            return
        
        # 顯示下拉選單
        self._show_dropdown()
    
    def _show_dropdown(self):
        """顯示下拉選單"""
        if self._state == "disabled":
            return
        
        # 創建下拉選單
        self._dropdown = ctk.CTkToplevel(self)
        self._dropdown.withdraw()  # 先隱藏
        self._dropdown.overrideredirect(True)  # 無邊框視窗
        self._dropdown.attributes("-topmost", True)
        
        # 設定 Toplevel 背景為透明（Windows）
        self._dropdown.wm_attributes("-transparentcolor", "gray1")
        self._dropdown.configure(fg_color="gray1")
        
        # 取得容器寬度（確保更新後再取得）
        self.container.update_idletasks()
        dropdown_width = self.container.winfo_width()
        
        # 圓角邊距（讓圓角完整顯示）
        corner_padding = 2
        
        # 下拉選單框架 - 使用 pack 並設定 padding
        dropdown_frame = ctk.CTkFrame(
            self._dropdown,
            fg_color="#ffffff",
            corner_radius=6,
            border_width=1,
            border_color="#cce5ff"
        )
        dropdown_frame.pack(fill="both", expand=True, padx=corner_padding, pady=corner_padding)
        
        # 選項（直接放在框架中，減少 padding）
        for i, value in enumerate(self._values):
            btn = ctk.CTkButton(
                dropdown_frame,
                text=value,
                font=Fonts.to_tuple(Fonts.BODY_SMALL),
                height=28,
                corner_radius=4,
                fg_color="transparent",
                hover_color="#e1f0ff",
                text_color="#37474f",
                anchor="w",
                command=lambda v=value: self._select_value(v)
            )
            btn.pack(fill="x", padx=4, pady=(4 if i == 0 else 1, 1))
        
        # 計算位置（往左上偏移以補償 padding）
        x = self.container.winfo_rootx() - corner_padding
        y = self.container.winfo_rooty() + self.container.winfo_height() + 2 - corner_padding
        
        # 設定下拉選單大小和位置（加上圓角邊距）
        dropdown_height = len(self._values) * 30 + 10 + corner_padding * 2
        total_width = dropdown_width + corner_padding * 2
        self._dropdown.geometry(f"{total_width}x{dropdown_height}+{x}+{y}")
        self._dropdown.deiconify()  # 顯示
        
        # 延遲綁定點擊外部事件
        self.after(100, self._bind_click_outside)
    
    def _bind_click_outside(self):
        """綁定點擊外部事件"""
        if self._dropdown and self._dropdown.winfo_exists():
            # 綁定到根視窗
            root = self.winfo_toplevel()
            self._click_binding = root.bind("<Button-1>", self._on_root_click, add="+")
    
    def _on_root_click(self, event):
        """處理根視窗點擊事件"""
        if not self._dropdown or not self._dropdown.winfo_exists():
            self._unbind_click()
            return
        
        # 取得點擊位置（螢幕座標）
        click_x, click_y = event.x_root, event.y_root
        
        # 檢查是否點擊在下拉選單內
        dx = self._dropdown.winfo_rootx()
        dy = self._dropdown.winfo_rooty()
        dw = self._dropdown.winfo_width()
        dh = self._dropdown.winfo_height()
        
        in_dropdown = dx <= click_x <= dx + dw and dy <= click_y <= dy + dh
        
        # 檢查是否點擊在 combo box 本身上
        cx = self.container.winfo_rootx()
        cy = self.container.winfo_rooty()
        cw = self.container.winfo_width()
        ch = self.container.winfo_height()
        
        in_combobox = cx <= click_x <= cx + cw and cy <= click_y <= cy + ch
        
        # 如果點擊在下拉選單或 combo box 外，則關閉
        if not in_dropdown and not in_combobox:
            self._close_dropdown()
    
    def _unbind_click(self):
        """解除點擊綁定"""
        if self._click_binding:
            try:
                self.winfo_toplevel().unbind("<Button-1>", self._click_binding)
            except:
                pass
            self._click_binding = None
    
    def _check_click_outside(self, event):
        """檢查是否點擊在下拉選單外"""
        if self._dropdown and self._dropdown.winfo_exists():
            # 檢查點擊位置
            x, y = event.x_root, event.y_root
            dx = self._dropdown.winfo_rootx()
            dy = self._dropdown.winfo_rooty()
            dw = self._dropdown.winfo_width()
            dh = self._dropdown.winfo_height()
            
            # 也檢查是否點擊在 combo box 本身上（用於切換）
            cx = self.container.winfo_rootx()
            cy = self.container.winfo_rooty()
            cw = self.container.winfo_width()
            ch = self.container.winfo_height()
            
            if cx <= x <= cx + cw and cy <= y <= cy + ch:
                # 點擊在 combo box 上，讓 _toggle_dropdown 處理
                return
            
            if not (dx <= x <= dx + dw and dy <= y <= dy + dh):
                self._close_dropdown()
    
    def _close_dropdown(self):
        """關閉下拉選單"""
        self._unbind_click()
        if self._dropdown and self._dropdown.winfo_exists():
            self._dropdown.destroy()
        self._dropdown = None
    
    def _select_value(self, value: str):
        """選擇值"""
        self._variable.set(value)
        self._close_dropdown()
        if self._command:
            self._command(value)
    
    def get(self) -> str:
        """取得當前值"""
        return self._variable.get()
    
    def set(self, value: str):
        """設定值"""
        self._variable.set(value)
    
    def configure(self, **kwargs):
        """設定屬性"""
        if "values" in kwargs:
            self._values = kwargs.pop("values")
        if "state" in kwargs:
            self._state = kwargs.pop("state")
            if self._state == "disabled":
                self.container.configure(fg_color="#f0f0f0")
                self.arrow_btn.configure(state="disabled")
            else:
                self.container.configure(fg_color="#ffffff")
                self.arrow_btn.configure(state="normal")
        if "variable" in kwargs:
            self._variable = kwargs.pop("variable")
            self.label.configure(textvariable=self._variable)
        super().configure(**kwargs)
    
    def cget(self, key: str):
        """取得屬性"""
        if key == "values":
            return self._values
        if key == "state":
            return self._state
        return super().cget(key)
