# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
驅動程式表格組件
支援多選、排序、搜尋的現代化表格
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from typing import Optional, Callable, Any, List, Dict
import re

from ui.theme import Fonts, theme_manager


def natural_sort_key(s: str) -> list:
    """
    自然排序 key 函數
    將字串中的數字部分轉換為整數進行比較
    例如: "oem10.inf" -> ["oem", 10, ".inf"]
    """
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split(r'(\d+)', s)]


class DriverTable(ctk.CTkFrame):
    """
    現代化驅動程式表格組件
    支援多選、排序、搜尋
    """
    
    # 欄位定義
    COLUMNS = ("select", "name", "inf", "provider", "version", "date", "class")
    COLUMN_WIDTHS = {
        "select": (40, 40, "center", False),
        "name": (180, 120, "w", True),
        "inf": (90, 70, "w", True),
        "provider": (150, 100, "w", True),
        "version": (100, 80, "w", True),
        "date": (100, 80, "w", True),
        "class": (100, 80, "w", True),
    }
    COLUMN_TITLES = {
        "select": "☐",
        "name": "驅動名稱",
        "inf": "INF 檔案",
        "provider": "提供者",
        "version": "版本",
        "date": "日期",
        "class": "類型",
    }
    # 欄位對應的資料 key
    DATA_KEY_MAP = {
        "name": "PublishedName",
        "inf": "OriginalFileName",
        "provider": "Provider",
        "version": "Version",
        "date": "Date",
        "class": "ClassName",
    }
    
    def __init__(
        self,
        master: Any,
        on_selection_change: Optional[Callable[[int], None]] = None,
        **kwargs
    ):
        """
        初始化驅動程式表格
        
        Args:
            master: 父組件
            on_selection_change: 選取變更回調 (selected_count)
        """
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self._on_selection_change = on_selection_change
        
        # 資料儲存
        self._drivers: List[Dict] = []
        self._filtered_drivers: List[Dict] = []
        self._selected_indices: set = set()
        self._sort_column: str = ""
        self._sort_reverse: bool = False
        self._search_text: str = ""
        self._select_all: bool = False
        self._tree_has_focus: bool = False
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        tree_container = ctk.CTkFrame(self, fg_color="transparent")
        tree_container.pack(fill="both", expand=True)
        
        # 設定 ttk 樣式
        self._setup_style()
        
        # 建立 Treeview
        self.tree = tk.ttk.Treeview(
            tree_container,
            columns=self.COLUMNS,
            show="headings",
            selectmode="extended",
            style="Driver.Treeview"
        )
        
        # 設定欄位
        self._setup_columns()
        
        # 綁定事件
        self._bind_events()
        
        # 滾動條
        scroll_y = ctk.CTkScrollbar(tree_container, command=self.tree.yview)
        scroll_x = ctk.CTkScrollbar(tree_container, orientation="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        # 佈局
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
    
    def _setup_style(self):
        """設定 ttk 樣式"""
        style = tk.ttk.Style()
        style.configure(
            "Driver.Treeview",
            background=theme_manager.colors.card_bg,
            foreground=theme_manager.colors.text_primary,
            fieldbackground=theme_manager.colors.card_bg,
            rowheight=28,
            font=Fonts.to_tuple(Fonts.BODY)
        )
        style.configure(
            "Driver.Treeview.Heading",
            background=theme_manager.colors.border,
            foreground=theme_manager.colors.text_primary,
            font=Fonts.to_tuple(Fonts.LABEL)
        )
        style.map(
            "Driver.Treeview", 
            background=[("selected", theme_manager.colors.primary)],
            foreground=[("selected", "#ffffff")]
        )
    
    def _setup_columns(self):
        """設定欄位"""
        for col in self.COLUMNS:
            title = self.COLUMN_TITLES.get(col, col)
            if col == "select":
                self.tree.heading(col, text=title, command=self._on_toggle_select_all)
            else:
                self.tree.heading(col, text=title, command=lambda c=col: self._on_sort(c))
            
            width, minwidth, anchor, stretch = self.COLUMN_WIDTHS.get(col, (100, 80, "w", True))
            self.tree.column(col, width=width, minwidth=minwidth, anchor=anchor, stretch=stretch)
    
    def _bind_events(self):
        """綁定事件"""
        self.tree.bind("<ButtonRelease-1>", self._on_row_click)
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<MouseWheel>", self._on_tree_mousewheel)
        self.tree.bind("<Enter>", self._on_tree_enter)
        self.tree.bind("<Leave>", self._on_tree_leave)
    
    # === 事件處理 ===
    
    def _on_toggle_select_all(self):
        """切換全選/取消全選"""
        self._select_all = not self._select_all
        self.tree.heading("select", text="☑" if self._select_all else "☐")
        
        if self._select_all:
            self._selected_indices = set(range(len(self._filtered_drivers)))
        else:
            self._selected_indices.clear()
        
        self._refresh_tree()
        self._notify_selection_change()
    
    def _on_row_click(self, event):
        """行點擊事件"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        if not item:
            return
        
        try:
            idx = self.tree.index(item)
        except:
            return
        
        # 只有點擊第一欄才切換選取
        if column == "#1":
            if idx in self._selected_indices:
                self._selected_indices.discard(idx)
            else:
                self._selected_indices.add(idx)
            
            self._refresh_row(item, idx)
            self._update_select_all_state()
            self._notify_selection_change()
    
    def _on_double_click(self, event):
        """雙擊查看詳情"""
        item = self.tree.identify_row(event.y)
        if item:
            try:
                idx = self.tree.index(item)
                if idx < len(self._filtered_drivers):
                    driver = self._filtered_drivers[idx]
                    self._show_driver_details(driver)
            except:
                pass
    
    def _on_tree_mousewheel(self, event):
        """處理 Treeview 的滾輪事件"""
        self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"
    
    def _on_tree_enter(self, event):
        """滑鼠進入 Treeview"""
        self._tree_has_focus = True
        self.winfo_toplevel().bind("<MouseWheel>", self._on_global_mousewheel)
    
    def _on_tree_leave(self, event):
        """滑鼠離開 Treeview"""
        self._tree_has_focus = False
        try:
            self.winfo_toplevel().unbind("<MouseWheel>")
        except:
            pass
    
    def _on_global_mousewheel(self, event):
        """全域滾輪事件處理"""
        if self._tree_has_focus:
            self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"
    
    def _show_driver_details(self, driver: Dict):
        """顯示驅動程式詳情"""
        details = "\n".join([f"{k}: {v}" for k, v in driver.items()])
        messagebox.showinfo("驅動程式詳情", details)
    
    def _on_sort(self, column: str):
        """排序"""
        if self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column = column
            self._sort_reverse = False
        
        self._apply_filter_and_sort()
    
    # === 內部方法 ===
    
    def _apply_filter_and_sort(self):
        """應用過濾和排序"""
        # 過濾
        if self._search_text:
            search_lower = self._search_text.lower()
            self._filtered_drivers = [
                d for d in self._drivers
                if any(search_lower in str(v).lower() for v in d.values())
            ]
        else:
            self._filtered_drivers = self._drivers.copy()
        
        # 排序
        if self._sort_column:
            key = self.DATA_KEY_MAP.get(self._sort_column, self._sort_column)
            self._filtered_drivers.sort(
                key=lambda d: natural_sort_key(str(d.get(key, ""))),
                reverse=self._sort_reverse
            )
        
        # 清空選取
        self._selected_indices.clear()
        self._select_all = False
        self.tree.heading("select", text="☐")
        
        self._refresh_tree()
        self._notify_selection_change()
    
    def _refresh_tree(self):
        """刷新表格顯示"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for i, driver in enumerate(self._filtered_drivers):
            selected = "☑" if i in self._selected_indices else "☐"
            values = (
                selected,
                driver.get("PublishedName", ""),
                driver.get("OriginalFileName", ""),
                driver.get("Provider", ""),
                driver.get("Version", ""),
                driver.get("Date", ""),
                driver.get("ClassName", "")
            )
            self.tree.insert("", "end", values=values)
    
    def _refresh_row(self, item: str, idx: int):
        """刷新單行"""
        if idx < len(self._filtered_drivers):
            driver = self._filtered_drivers[idx]
            selected = "☑" if idx in self._selected_indices else "☐"
            values = (
                selected,
                driver.get("PublishedName", ""),
                driver.get("OriginalFileName", ""),
                driver.get("Provider", ""),
                driver.get("Version", ""),
                driver.get("Date", ""),
                driver.get("ClassName", "")
            )
            self.tree.item(item, values=values)
    
    def _update_select_all_state(self):
        """更新全選狀態"""
        total = len(self._filtered_drivers)
        selected = len(self._selected_indices)
        
        if selected == total and total > 0:
            self._select_all = True
            self.tree.heading("select", text="☑")
        else:
            self._select_all = False
            self.tree.heading("select", text="☐")
    
    def _notify_selection_change(self):
        """通知選取變更"""
        if self._on_selection_change:
            self._on_selection_change(len(self._selected_indices))
    
    # === 公開方法 ===
    
    def set_drivers(self, drivers: List[Dict]):
        """設定驅動程式列表"""
        self._drivers = drivers
        self._search_text = ""
        self._selected_indices.clear()
        self._apply_filter_and_sort()
    
    def search(self, text: str):
        """搜尋"""
        self._search_text = text.strip()
        self._apply_filter_and_sort()
    
    def clear_search(self):
        """清除搜尋"""
        self._search_text = ""
        self._apply_filter_and_sort()
    
    def get_selected_drivers(self) -> List[Dict]:
        """取得已選取的驅動程式"""
        return [
            self._filtered_drivers[i] 
            for i in sorted(self._selected_indices)
            if i < len(self._filtered_drivers)
        ]
    
    def get_selected_published_names(self) -> List[str]:
        """取得已選取的 PublishedName 列表"""
        return [d.get("PublishedName", "") for d in self.get_selected_drivers()]
    
    def get_count(self) -> tuple:
        """取得 (總數, 已選取數)"""
        return len(self._filtered_drivers), len(self._selected_indices)
    
    def clear(self):
        """清空表格"""
        self._drivers.clear()
        self._filtered_drivers.clear()
        self._selected_indices.clear()
        self._refresh_tree()
