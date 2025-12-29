# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
驅動程式管理頁面
現代化 UI 設計
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional, Callable, Any, List, Dict
import os
import threading

from ui.theme import ThemeManager, Fonts, Spacing, theme_manager
from ui.components import (
    ModernButton, ModernCard, ModernEntry, ModernTooltip, 
    StatusBadge, FormField, SectionTitle, ModernProgressBar
)
from app.driver_manager import DriverManager


class DriverTable(ctk.CTkFrame):
    """
    現代化驅動程式表格組件
    支援多選、排序、搜尋
    """
    
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
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # 表格容器（使用原生 ttk.Treeview 因為 CustomTkinter 沒有表格組件）
        tree_container = ctk.CTkFrame(self, fg_color="transparent")
        tree_container.pack(fill="both", expand=True)
        
        # 設定 ttk 樣式
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
        style.map("Driver.Treeview", 
            background=[("selected", theme_manager.colors.primary)],
            foreground=[("selected", "#ffffff")]
        )
        
        # 建立 Treeview
        columns = ("select", "name", "inf", "provider", "version", "date", "class")
        self.tree = tk.ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",
            selectmode="extended",
            style="Driver.Treeview"
        )
        
        # 設定欄位標題
        self._select_all = False
        self.tree.heading("select", text="☐", command=self._on_toggle_select_all)
        self.tree.heading("name", text="驅動名稱", command=lambda: self._on_sort("name"))
        self.tree.heading("inf", text="INF 檔案", command=lambda: self._on_sort("inf"))
        self.tree.heading("provider", text="提供者", command=lambda: self._on_sort("provider"))
        self.tree.heading("version", text="版本", command=lambda: self._on_sort("version"))
        self.tree.heading("date", text="日期", command=lambda: self._on_sort("date"))
        self.tree.heading("class", text="類型", command=lambda: self._on_sort("class"))
        
        # 設定欄位寬度
        self.tree.column("select", width=40, minwidth=40, anchor="center", stretch=False)
        self.tree.column("name", width=180, minwidth=120)
        self.tree.column("inf", width=90, minwidth=70)
        self.tree.column("provider", width=150, minwidth=100)
        self.tree.column("version", width=100, minwidth=80)
        self.tree.column("date", width=100, minwidth=80)
        self.tree.column("class", width=100, minwidth=80)
        
        # 點擊事件
        self.tree.bind("<ButtonRelease-1>", self._on_row_click)
        self.tree.bind("<Double-1>", self._on_double_click)
        
        # 滾輪事件 - 讓 Treeview 自己處理滾輪，不傳播到父視窗
        self.tree.bind("<MouseWheel>", self._on_tree_mousewheel)
        self.tree.bind("<Enter>", self._on_tree_enter)
        self.tree.bind("<Leave>", self._on_tree_leave)
        self._tree_has_focus = False
        
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
    
    def _on_toggle_select_all(self):
        """切換全選/取消全選"""
        self._select_all = not self._select_all
        
        # 更新標題
        self.tree.heading("select", text="☑" if self._select_all else "☐")
        
        # 更新所有行
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
        
        # 取得行索引
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
        # 滾動 Treeview
        self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
        # 返回 "break" 阻止事件繼續傳播到父視窗
        return "break"
    
    def _on_tree_enter(self, event):
        """滑鼠進入 Treeview"""
        self._tree_has_focus = True
        # 綁定全域滾輪事件到 Treeview
        self.winfo_toplevel().bind("<MouseWheel>", self._on_global_mousewheel)
    
    def _on_tree_leave(self, event):
        """滑鼠離開 Treeview"""
        self._tree_has_focus = False
        # 解除綁定
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
            key_map = {
                "name": "OriginalFileName",
                "inf": "PublishedName",
                "provider": "Provider",
                "version": "Version",
                "date": "Date",
                "class": "ClassName"
            }
            key = key_map.get(self._sort_column, self._sort_column)
            self._filtered_drivers.sort(
                key=lambda d: str(d.get(key, "")).lower(),
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
        # 清空
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 重新填入
        for i, driver in enumerate(self._filtered_drivers):
            selected = "☑" if i in self._selected_indices else "☐"
            values = (
                selected,
                driver.get("OriginalFileName", ""),
                driver.get("PublishedName", ""),
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
                driver.get("OriginalFileName", ""),
                driver.get("PublishedName", ""),
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


class DriverPage(ctk.CTkFrame):
    """
    驅動程式管理頁面
    """
    
    def __init__(
        self,
        master: Any,
        on_log: Callable[[str], None],
        get_mounted_dirs: Callable[[], List[tuple]],
        show_header: bool = True,
        **kwargs
    ):
        """
        初始化驅動程式頁面
        
        Args:
            master: 父組件
            on_log: 日誌回調
            get_mounted_dirs: 取得已掛載目錄的回調 [("WIM#1", path), ...]
            show_header: 是否顯示頁面標題
        """
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self._on_log = on_log
        self._get_mounted_dirs = get_mounted_dirs
        self._show_header = show_header
        
        # 變數
        self.var_target_mount = ctk.StringVar()
        self.var_search = ctk.StringVar()
        self.var_status = ctk.StringVar(value="請選擇已掛載的映像")
        
        self._build_ui()
    
    def _build_ui(self):
        """建立 UI"""
        # 頁面標題（可選）
        if self._show_header:
            header = ctk.CTkFrame(self, fg_color="transparent")
            header.pack(fill="x", padx=20, pady=(20, 16))
            
            ctk.CTkLabel(
                header,
                text="驅動程式管理",
                font=Fonts.to_tuple(Fonts.TITLE),
                text_color=theme_manager.colors.text_primary
            ).pack(side="left")
            
            ctk.CTkLabel(
                header,
                text="安裝、移除、萃取映像中的驅動程式",
                font=Fonts.to_tuple(Fonts.BODY),
                text_color=theme_manager.colors.text_muted
            ).pack(side="left", padx=(12, 0))
        
        # 根據是否顯示標題來調整 padding
        pad_x = 20 if self._show_header else 0
        pad_y = (0, 12) if self._show_header else (0, 8)
        
        # === 目標映像選擇區 ===
        target_card = ModernCard(self, padding=12)
        target_card.pack(fill="x", padx=pad_x, pady=pad_y)
        target_content = target_card.get_content_frame()
        
        target_row = ctk.CTkFrame(target_content, fg_color="transparent")
        target_row.pack(fill="x")
        
        ctk.CTkLabel(
            target_row,
            text="目標映像",
            font=Fonts.to_tuple(Fonts.LABEL),
            text_color=theme_manager.colors.text_secondary,
            width=80,
            anchor="w"
        ).pack(side="left")
        
        self.combo_target = ctk.CTkComboBox(
            target_row,
            width=400,
            variable=self.var_target_mount,
            state="readonly",
            values=[""],
            font=Fonts.to_tuple(Fonts.BODY),
            dropdown_font=Fonts.to_tuple(Fonts.BODY),
            command=self._on_target_changed
        )
        self.combo_target.pack(side="left", padx=(8, 0))
        
        self.status_badge = StatusBadge(target_row, "未選擇", "default")
        self.status_badge.pack(side="left", padx=(12, 0))
        
        # 重新整理按鈕
        ModernButton(
            target_row,
            text="重新整理",
            variant="outline",
            size="sm",
            command=self._refresh_driver_list
        ).pack(side="left", padx=(12, 0))
        
        # 搜尋
        search_frame = ctk.CTkFrame(target_row, fg_color="transparent")
        search_frame.pack(side="right")
        
        ctk.CTkLabel(
            search_frame,
            text="🔍",
            font=Fonts.to_tuple(Fonts.BODY)
        ).pack(side="left")
        
        self.entry_search = ModernEntry(
            search_frame,
            placeholder="搜尋驅動...",
            width=180
        )
        self.entry_search.pack(side="left", padx=(4, 0))
        self.entry_search.configure(textvariable=self.var_search)
        self.entry_search.bind("<Return>", lambda e: self._on_search())
        self.entry_search.bind("<Escape>", lambda e: self._on_clear_search())
        
        ModernButton(
            search_frame,
            text="搜尋",
            variant="ghost",
            size="sm",
            command=self._on_search
        ).pack(side="left", padx=(4, 0))
        
        # === 驅動程式表格 ===
        table_frame = ctk.CTkFrame(self, fg_color="transparent")
        table_frame.pack(fill="both", expand=True, padx=20, pady=(0, 12))
        
        self.driver_table = DriverTable(
            table_frame,
            on_selection_change=self._on_selection_change
        )
        self.driver_table.pack(fill="both", expand=True)
        
        # === 狀態列 ===
        status_frame = ctk.CTkFrame(self, fg_color="transparent")
        status_frame.pack(fill="x", padx=20, pady=(0, 8))
        
        self.lbl_status = ctk.CTkLabel(
            status_frame,
            textvariable=self.var_status,
            font=Fonts.to_tuple(Fonts.CAPTION),
            text_color=theme_manager.colors.text_muted
        )
        self.lbl_status.pack(side="left")
        
        # === 操作按鈕區 ===
        btn_card = ModernCard(self, padding=12)
        btn_card.pack(fill="x", padx=20, pady=(0, 20))
        btn_content = btn_card.get_content_frame()
        
        btn_row = ctk.CTkFrame(btn_content, fg_color="transparent")
        btn_row.pack(fill="x")
        
        # 左側主要操作
        left_btns = ctk.CTkFrame(btn_row, fg_color="transparent")
        left_btns.pack(side="left")
        
        self.btn_extract = ModernButton(
            left_btns,
            text="📤 萃取驅動",
            variant="primary",
            command=self._on_extract_drivers
        )
        self.btn_extract.pack(side="left")
        ModernTooltip(self.btn_extract, "將選取的驅動程式萃取到指定目錄")
        
        self.btn_add = ModernButton(
            left_btns,
            text="➕ 新增驅動",
            variant="secondary",
            command=self._on_add_driver
        )
        self.btn_add.pack(side="left", padx=(12, 0))
        ModernTooltip(self.btn_add, "安裝新的驅動程式到映像中")
        
        self.btn_remove = ModernButton(
            left_btns,
            text="🗑️ 移除驅動",
            variant="danger",
            command=self._on_remove_drivers
        )
        self.btn_remove.pack(side="left", padx=(12, 0))
        ModernTooltip(self.btn_remove, "從映像中移除選取的驅動程式")
        
        # 右側輔助操作
        right_btns = ctk.CTkFrame(btn_row, fg_color="transparent")
        right_btns.pack(side="right")
        
        ModernButton(
            right_btns,
            text="查看詳情",
            variant="ghost",
            size="sm",
            command=self._on_view_details
        ).pack(side="left")
        
        ModernButton(
            right_btns,
            text="匯出清單",
            variant="ghost",
            size="sm",
            command=self._on_export_list
        ).pack(side="left", padx=(8, 0))
        
        # 初始化按鈕狀態
        self._update_button_states()
    
    # === 事件處理 ===
    
    def _on_target_changed(self, selection: str):
        """目標映像選擇變更"""
        # 解析選擇
        mount_dir = ""
        if selection:
            # 格式: "WIM#1 - path (status)" 或類似
            parts = selection.split(" - ")
            if len(parts) >= 2:
                path_part = parts[1].split(" (")[0]
                mount_dir = path_part.strip()
        
        if mount_dir and os.path.isdir(mount_dir):
            # 檢查掛載狀態
            from app.wim_manager import WIMManager
            is_mounted, _, _ = WIMManager.is_path_mounted(mount_dir)
            
            if is_mounted:
                self.status_badge.set_status("已掛載", "success")
                self._refresh_driver_list()
            else:
                self.status_badge.set_status("未掛載", "warning")
                self.driver_table.clear()
                self.var_status.set("映像未掛載，請先掛載後再操作")
        else:
            self.status_badge.set_status("未選擇", "default")
            self.driver_table.clear()
            self.var_status.set("請選擇已掛載的映像")
        
        self._update_button_states()
    
    def _on_selection_change(self, selected_count: int):
        """選取變更"""
        total, selected = self.driver_table.get_count()
        self.var_status.set(f"共 {total} 個驅動程式，已選取 {selected} 個")
    
    def _on_search(self):
        """執行搜尋"""
        text = self.var_search.get().strip()
        self.driver_table.search(text)
        if text:
            self._on_log(f"搜尋: {text}")
    
    def _on_clear_search(self):
        """清除搜尋"""
        self.var_search.set("")
        self.driver_table.clear_search()
    
    def _refresh_driver_list(self):
        """重新整理驅動程式列表"""
        mount_dir = self._get_current_mount_dir()
        if not mount_dir:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        self._on_log("正在讀取驅動程式列表...")
        self.var_status.set("正在讀取...")
        
        def do_refresh():
            success, drivers, error = DriverManager.get_drivers_in_offline_image(mount_dir)
            self.after(0, lambda: self._refresh_complete(success, drivers, error))
        
        threading.Thread(target=do_refresh, daemon=True).start()
    
    def _refresh_complete(self, success: bool, drivers: List[Dict], error: str):
        """重新整理完成"""
        if success:
            self.driver_table.set_drivers(drivers)
            total, _ = self.driver_table.get_count()
            self.var_status.set(f"共 {total} 個驅動程式")
            self._on_log(f"✓ 已載入 {total} 個驅動程式")
        else:
            self.driver_table.clear()
            self.var_status.set(f"讀取失敗: {error}")
            self._on_log(f"✗ 讀取驅動程式失敗: {error}")
        
        self._update_button_states()
    
    def _on_extract_drivers(self):
        """萃取驅動程式"""
        mount_dir = self._get_current_mount_dir()
        if not mount_dir:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        selected = self.driver_table.get_selected_drivers()
        selected_names = self.driver_table.get_selected_published_names()
        
        # 選擇輸出目錄
        output_dir = filedialog.askdirectory(title="選擇萃取輸出目錄")
        if not output_dir:
            return
        
        self.btn_extract.set_loading(True)
        
        def do_extract():
            if not selected:
                # 萃取全部
                self._on_log(f"正在萃取所有驅動程式到: {output_dir}")
                success, msg = DriverManager.export_drivers_from_offline_image(mount_dir, output_dir)
                self.after(0, lambda: self._extract_complete(success, msg, output_dir))
            else:
                # 萃取選定的驅動
                self._on_log(f"正在萃取 {len(selected)} 個驅動程式到: {output_dir}")
                
                def callback(current, total, name, success, msg):
                    status = "✓" if success else "✗"
                    self.after(0, lambda: self._on_log(f"  {status} [{current}/{total}] {name}"))
                
                success_count, fail_count, errors = DriverManager.export_selected_drivers(
                    mount_dir, output_dir, selected_names, callback
                )
                self.after(0, lambda: self._extract_selected_complete(success_count, fail_count, errors, output_dir))
        
        threading.Thread(target=do_extract, daemon=True).start()
    
    def _extract_complete(self, success: bool, message: str, output_dir: str):
        """萃取完成"""
        self.btn_extract.set_loading(False)
        
        if success:
            self._on_log(f"✓ 驅動程式萃取完成")
            if messagebox.askyesno("完成", "驅動程式萃取完成！\n是否開啟輸出目錄？"):
                os.startfile(output_dir)
        else:
            self._on_log(f"✗ 萃取失敗: {message}")
            messagebox.showerror("錯誤", f"萃取失敗: {message}")
    
    def _extract_selected_complete(self, success_count: int, fail_count: int, errors: List[str], output_dir: str):
        """選擇性萃取完成"""
        self.btn_extract.set_loading(False)
        
        self._on_log(f"萃取完成: 成功 {success_count}，失敗 {fail_count}")
        
        if fail_count > 0:
            error_msg = "\n".join(errors[:5])
            if len(errors) > 5:
                error_msg += f"\n... 還有 {len(errors) - 5} 個錯誤"
            messagebox.showwarning("部分失敗", f"成功: {success_count}\n失敗: {fail_count}\n\n{error_msg}")
        else:
            if messagebox.askyesno("完成", f"已成功萃取 {success_count} 個驅動程式！\n是否開啟輸出目錄？"):
                os.startfile(output_dir)
    
    def _on_add_driver(self):
        """新增驅動程式"""
        mount_dir = self._get_current_mount_dir()
        if not mount_dir:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        # 彈出新增驅動對話框
        self._show_add_driver_dialog(mount_dir)
    
    def _show_add_driver_dialog(self, mount_dir: str):
        """顯示新增驅動對話框"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("新增驅動程式")
        dialog.geometry("500x280")
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        
        # 居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - 500) // 2
        y = (dialog.winfo_screenheight() - 280) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 內容
        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(
            content,
            text="新增驅動程式",
            font=Fonts.to_tuple(Fonts.TITLE_SMALL)
        ).pack(anchor="w", pady=(0, 16))
        
        # 來源路徑
        var_source = ctk.StringVar()
        
        source_frame = ctk.CTkFrame(content, fg_color="transparent")
        source_frame.pack(fill="x", pady=(0, 12))
        
        ctk.CTkLabel(
            source_frame,
            text="驅動來源",
            font=Fonts.to_tuple(Fonts.LABEL),
            width=80,
            anchor="w"
        ).pack(side="left")
        
        entry_source = ModernEntry(source_frame, width=280)
        entry_source.pack(side="left", padx=(8, 0))
        entry_source.configure(textvariable=var_source)
        
        def browse_file():
            path = filedialog.askopenfilename(
                title="選擇驅動程式",
                filetypes=[("INF 檔案", "*.inf"), ("所有檔案", "*.*")]
            )
            if path:
                var_source.set(path)
        
        def browse_folder():
            path = filedialog.askdirectory(title="選擇驅動資料夾")
            if path:
                var_source.set(path)
        
        ModernButton(
            source_frame,
            text="檔案",
            variant="outline",
            size="sm",
            command=browse_file
        ).pack(side="left", padx=(8, 0))
        
        ModernButton(
            source_frame,
            text="資料夾",
            variant="outline",
            size="sm",
            command=browse_folder
        ).pack(side="left", padx=(4, 0))
        
        # 選項
        var_recurse = ctk.BooleanVar(value=True)
        var_force = ctk.BooleanVar(value=False)
        
        opt_frame = ctk.CTkFrame(content, fg_color="transparent")
        opt_frame.pack(fill="x", pady=(0, 16))
        
        ctk.CTkCheckBox(
            opt_frame,
            text="遞迴搜尋子資料夾",
            variable=var_recurse,
            font=Fonts.to_tuple(Fonts.BODY)
        ).pack(side="left")
        
        ctk.CTkCheckBox(
            opt_frame,
            text="強制安裝未簽署驅動",
            variable=var_force,
            font=Fonts.to_tuple(Fonts.BODY)
        ).pack(side="left", padx=(24, 0))
        
        # 按鈕
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(16, 0))
        
        def do_add():
            source = var_source.get().strip()
            if not source:
                messagebox.showwarning("提示", "請選擇驅動程式來源")
                return
            
            self._on_log(f"正在安裝驅動: {source}")
            
            def add_thread():
                success, msg = DriverManager.add_driver_to_offline_image(
                    mount_dir, source,
                    recurse=var_recurse.get(),
                    force_unsigned=var_force.get()
                )
                self.after(0, lambda: add_complete(success, msg))
            
            def add_complete(success, msg):
                if success:
                    self._on_log("✓ 驅動程式安裝完成")
                    messagebox.showinfo("完成", "驅動程式安裝完成！")
                    dialog.destroy()
                    self._refresh_driver_list()
                else:
                    self._on_log(f"✗ 安裝失敗: {msg}")
                    messagebox.showerror("錯誤", f"安裝失敗: {msg}")
            
            threading.Thread(target=add_thread, daemon=True).start()
        
        ModernButton(
            btn_frame,
            text="取消",
            variant="ghost",
            command=dialog.destroy
        ).pack(side="right")
        
        ModernButton(
            btn_frame,
            text="安裝",
            variant="primary",
            command=do_add
        ).pack(side="right", padx=(0, 12))
    
    def _on_remove_drivers(self):
        """移除驅動程式"""
        mount_dir = self._get_current_mount_dir()
        if not mount_dir:
            messagebox.showwarning("提示", "請先選擇目標映像")
            return
        
        selected = self.driver_table.get_selected_published_names()
        if not selected:
            messagebox.showwarning("提示", "請先選取要移除的驅動程式")
            return
        
        if not messagebox.askyesno("確認", f"確定要移除 {len(selected)} 個驅動程式嗎？\n此操作無法復原！"):
            return
        
        self._on_log(f"正在移除 {len(selected)} 個驅動程式...")
        self.btn_remove.set_loading(True)
        
        def do_remove():
            def callback(current, total, name, success, msg):
                status = "✓" if success else "✗"
                self.after(0, lambda: self._on_log(f"  {status} [{current}/{total}] {name}"))
            
            success_count, fail_count, errors = DriverManager.remove_drivers_batch(
                mount_dir, selected, callback
            )
            self.after(0, lambda: self._remove_complete(success_count, fail_count, errors))
        
        threading.Thread(target=do_remove, daemon=True).start()
    
    def _remove_complete(self, success_count: int, fail_count: int, errors: List[str]):
        """移除完成"""
        self.btn_remove.set_loading(False)
        
        self._on_log(f"移除完成: 成功 {success_count}，失敗 {fail_count}")
        
        if fail_count > 0:
            error_msg = "\n".join(errors[:5])
            if len(errors) > 5:
                error_msg += f"\n... 還有 {len(errors) - 5} 個錯誤"
            messagebox.showwarning("部分失敗", f"成功: {success_count}\n失敗: {fail_count}\n\n{error_msg}")
        else:
            messagebox.showinfo("完成", f"已成功移除 {success_count} 個驅動程式")
        
        self._refresh_driver_list()
    
    def _on_view_details(self):
        """查看詳情"""
        selected = self.driver_table.get_selected_drivers()
        if not selected:
            messagebox.showinfo("提示", "請先選取驅動程式")
            return
        
        # 顯示第一個選取的詳情
        driver = selected[0]
        details = "\n".join([f"{k}: {v}" for k, v in driver.items()])
        messagebox.showinfo("驅動程式詳情", details)
    
    def _on_export_list(self):
        """匯出清單"""
        total, _ = self.driver_table.get_count()
        if total == 0:
            messagebox.showinfo("提示", "沒有驅動程式可匯出")
            return
        
        path = filedialog.asksaveasfilename(
            title="匯出驅動清單",
            defaultextension=".txt",
            filetypes=[("文字檔", "*.txt"), ("CSV", "*.csv")]
        )
        if not path:
            return
        
        try:
            with open(path, "w", encoding="utf-8") as f:
                drivers = self.driver_table._filtered_drivers
                f.write("PublishedName\tOriginalFileName\tProvider\tVersion\tDate\tClassName\n")
                for d in drivers:
                    f.write(f"{d.get('PublishedName', '')}\t")
                    f.write(f"{d.get('OriginalFileName', '')}\t")
                    f.write(f"{d.get('Provider', '')}\t")
                    f.write(f"{d.get('Version', '')}\t")
                    f.write(f"{d.get('Date', '')}\t")
                    f.write(f"{d.get('ClassName', '')}\n")
            
            self._on_log(f"✓ 驅動清單已匯出至: {path}")
            messagebox.showinfo("完成", "清單匯出成功！")
        except Exception as e:
            self._on_log(f"✗ 匯出失敗: {e}")
            messagebox.showerror("錯誤", f"匯出失敗: {e}")
    
    def _get_current_mount_dir(self) -> str:
        """取得當前選擇的掛載目錄"""
        selection = self.var_target_mount.get()
        if selection:
            parts = selection.split(" - ")
            if len(parts) >= 2:
                path_part = parts[1].split(" (")[0]
                return path_part.strip()
        return ""
    
    def _update_button_states(self):
        """更新按鈕狀態"""
        mount_dir = self._get_current_mount_dir()
        has_mount = bool(mount_dir and os.path.isdir(mount_dir))
        
        # 檢查是否已掛載
        is_mounted = False
        if has_mount:
            from app.wim_manager import WIMManager
            is_mounted, _, _ = WIMManager.is_path_mounted(mount_dir)
        
        state = "normal" if is_mounted else "disabled"
        self.btn_extract.configure(state=state)
        self.btn_add.configure(state=state)
        self.btn_remove.configure(state=state)
    
    # === 公開方法 ===
    
    def refresh_targets(self):
        """重新整理目標映像選項"""
        mounted_dirs = self._get_mounted_dirs()
        
        from app.wim_manager import WIMManager
        
        options = []
        for name, path in mounted_dirs:
            if path:
                is_mounted, _, _ = WIMManager.is_path_mounted(path)
                status = "✓ 已掛載" if is_mounted else "○ 未掛載"
                options.append(f"{name} - {path} ({status})")
            else:
                options.append(f"{name} - (未設定)")
        
        if not options:
            options = ["(無可用映像)"]
        
        self.combo_target.configure(values=options)
        
        # 自動選擇第一個已掛載的
        for i, opt in enumerate(options):
            if "✓ 已掛載" in opt:
                self.combo_target.set(opt)
                self._on_target_changed(opt)
                break
        else:
            if options:
                self.combo_target.set(options[0])
                self._on_target_changed(options[0])
