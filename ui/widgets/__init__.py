# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UI Widgets 模組
可重用的 UI 小部件
"""

# 驅動表格
from ui.widgets.driver_table import DriverTable, natural_sort_key

# 按鈕組件
from ui.widgets.button import ModernButton, IconButton

# 卡片組件
from ui.widgets.card import ModernCard, SectionTitle, EmptyState

# 表單組件
from ui.widgets.form import ModernEntry, FormField, ModernComboBox

# 回饋組件
from ui.widgets.feedback import (
    ModernTooltip,
    StatusBadge,
    ModernProgressBar,
    ModernSwitch,
)

# 對話框組件
from ui.widgets.dialog import (
    ModernDialog,
    ConfirmDialog,
    UnmountWarningDialog,
)

__all__ = [
    # 驅動表格
    'DriverTable',
    'natural_sort_key',
    # 按鈕
    'ModernButton',
    'IconButton',
    # 卡片
    'ModernCard',
    'SectionTitle',
    'EmptyState',
    # 表單
    'ModernEntry',
    'FormField',
    'ModernComboBox',
    # 回饋
    'ModernTooltip',
    'StatusBadge',
    'ModernProgressBar',
    'ModernSwitch',
    # 對話框
    'ModernDialog',
    'ConfirmDialog',
    'UnmountWarningDialog',
]
