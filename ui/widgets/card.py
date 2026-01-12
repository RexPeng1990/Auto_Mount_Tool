# -*- coding: utf-8 -*-
"""
卡片與區段組件模組

提供 ModernCard, SectionTitle, EmptyState 組件
"""

import customtkinter as ctk
from typing import Optional, Any
from ui.theme import Fonts, theme_manager


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
        
        # 右側：操作按鈕 (延遲匯入避免循環依賴)
        if action_button:
            from ui.widgets.button import ModernButton
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
        
        # 操作按鈕 (延遲匯入避免循環依賴)
        if action_button:
            from ui.widgets.button import ModernButton
            btn = ModernButton(
                self,
                text=action_button.get("text", ""),
                command=action_button.get("command"),
                variant=action_button.get("variant", "primary"),
                icon=action_button.get("icon")
            )
            btn.pack(pady=(16, 0))
