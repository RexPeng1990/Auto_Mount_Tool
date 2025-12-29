# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
現代化 UI 模組
使用 CustomTkinter 框架
"""

from ui.theme import ThemeManager, Colors, Fonts, theme_manager
from ui.components import (
    ModernButton,
    ModernEntry,
    ModernCard,
    ModernTooltip,
    StatusBadge,
    IconButton,
    ModernProgressBar,
    ModernSwitch,
    FormField,
    SectionTitle,
    EmptyState,
)
from ui.log_panel import LogPanel, CollapsibleLogPanel
from ui.pages.wim_page import WIMPage, WIMSlot
from ui.pages.driver_page import DriverPage, DriverTable

__all__ = [
    # Theme
    'ThemeManager',
    'Colors',
    'Fonts',
    'theme_manager',
    # Components
    'ModernButton',
    'ModernEntry',
    'ModernCard',
    'ModernTooltip',
    'StatusBadge',
    'IconButton',
    'ModernProgressBar',
    'ModernSwitch',
    'FormField',
    'SectionTitle',
    'EmptyState',
    # Log
    'LogPanel',
    'CollapsibleLogPanel',
    # Pages
    'WIMPage',
    'WIMSlot',
    'DriverPage',
    'DriverTable',
]
