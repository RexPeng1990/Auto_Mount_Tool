# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
單元測試套件 - 測試主程式模組

測試範圍:
1. app/config.py - 設定管理
2. app/utils.py - 工具函數
3. ui/theme.py - 主題系統
4. ui/collapsible.py - 折疊組件
5. ui/components.py - UI 組件
6. main_modern.py - 主程式整合
"""

import unittest
import sys
import os

# 確保專案根目錄在路徑中
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)
os.chdir(PROJECT_ROOT)


class TestConfigModule(unittest.TestCase):
    """測試設定模組"""
    
    def test_config_imports(self):
        """測試設定模組可正常匯入"""
        from app.config import CONFIG_FILE, LOG_DIR, DRIVER_EXPORT_DIR
        self.assertIsInstance(CONFIG_FILE, str)
        self.assertIsInstance(LOG_DIR, str)
        self.assertIsInstance(DRIVER_EXPORT_DIR, str)
    
    def test_config_paths_exist(self):
        """測試路徑常數為有效路徑格式"""
        from app.config import CONFIG_FILE, LOG_DIR
        # 檢查路徑格式（不需要實際存在）
        self.assertTrue(CONFIG_FILE.endswith('.ini'))
        self.assertIn('log', LOG_DIR.lower())
    
    def test_config_manager_creation(self):
        """測試 ConfigManager 可正常建立"""
        from app.config import ConfigManager
        cm = ConfigManager()
        self.assertIsNotNone(cm)
        self.assertIsNotNone(cm.cfg)
    
    def test_config_manager_get_default(self):
        """測試 ConfigManager.get 的預設值"""
        from app.config import ConfigManager
        cm = ConfigManager()
        # 測試不存在的設定應返回 fallback
        result = cm.get('NONEXISTENT', 'option', fallback='default_value')
        self.assertEqual(result, 'default_value')
    
    def test_config_manager_get_bool(self):
        """測試 ConfigManager.get_bool"""
        from app.config import ConfigManager
        cm = ConfigManager()
        # 測試不存在的設定應返回 fallback
        result = cm.get_bool('NONEXISTENT', 'option', fallback=True)
        self.assertTrue(result)
        result = cm.get_bool('NONEXISTENT', 'option', fallback=False)
        self.assertFalse(result)


class TestThemeModule(unittest.TestCase):
    """測試主題模組"""
    
    def test_theme_imports(self):
        """測試主題模組可正常匯入"""
        from ui.theme import theme_manager, ThemeManager, Colors, ColorScheme, Fonts, Spacing
        self.assertIsNotNone(theme_manager)
        self.assertIsNotNone(Colors)
        self.assertIsNotNone(Fonts)
    
    def test_color_scheme_properties(self):
        """測試 ColorScheme 的所有屬性都可存取"""
        from ui.theme import Colors
        
        # 測試深色主題
        dark = Colors.DARK
        required_attrs = [
            'primary', 'primary_hover', 'primary_disabled',
            'secondary', 'secondary_hover',
            'success', 'warning', 'danger', 'info',
            'bg_primary', 'bg_secondary', 'bg_tertiary', 'bg_card', 'bg_hover',
            'text_primary', 'text_secondary', 'text_muted', 'text_inverse',
            'border', 'border_hover', 'border_focus',
            'background', 'card_bg'  # 別名屬性
        ]
        
        for attr in required_attrs:
            with self.subTest(attr=attr):
                value = getattr(dark, attr)
                self.assertIsInstance(value, str, f"{attr} should be a string")
                self.assertTrue(value.startswith('#') or value.startswith('rgb'), 
                              f"{attr} should be a color value")
    
    def test_color_scheme_aliases(self):
        """測試 ColorScheme 別名屬性正確映射"""
        from ui.theme import Colors
        
        dark = Colors.DARK
        self.assertEqual(dark.background, dark.bg_primary)
        self.assertEqual(dark.card_bg, dark.bg_card)
        
        light = Colors.LIGHT
        self.assertEqual(light.background, light.bg_primary)
        self.assertEqual(light.card_bg, light.bg_card)
    
    def test_fonts_all_defined(self):
        """測試所有字體配置都已定義"""
        from ui.theme import Fonts, FontConfig
        
        required_fonts = [
            'TITLE_LARGE', 'TITLE', 'TITLE_SMALL',
            'BODY_LARGE', 'BODY', 'BODY_SMALL',
            'CAPTION', 'CAPTION_SMALL',
            'BUTTON', 'BUTTON_SMALL',
            'CODE', 'CODE_SMALL',
            'LABEL', 'LABEL_SMALL'
        ]
        
        for font_name in required_fonts:
            with self.subTest(font=font_name):
                font_config = getattr(Fonts, font_name, None)
                self.assertIsNotNone(font_config, f"Font {font_name} should be defined")
                self.assertIsInstance(font_config, FontConfig)
    
    def test_fonts_to_tuple(self):
        """測試 Fonts.to_tuple 正確轉換"""
        from ui.theme import Fonts
        
        result = Fonts.to_tuple(Fonts.TITLE)
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 3)
        self.assertIsInstance(result[0], str)  # family
        self.assertIsInstance(result[1], int)  # size
        self.assertIsInstance(result[2], str)  # weight
    
    def test_theme_manager_singleton(self):
        """測試 theme_manager 是單例"""
        from ui.theme import theme_manager, ThemeManager
        self.assertIsInstance(theme_manager, ThemeManager)
    
    def test_theme_manager_colors_property(self):
        """測試 theme_manager.colors 屬性"""
        from ui.theme import theme_manager, ColorScheme
        colors = theme_manager.colors
        self.assertIsInstance(colors, ColorScheme)
    
    def test_theme_manager_button_colors(self):
        """測試 theme_manager.get_button_colors"""
        from ui.theme import theme_manager
        
        variants = ['primary', 'secondary', 'success', 'warning', 'danger', 'outline', 'ghost']
        for variant in variants:
            with self.subTest(variant=variant):
                colors = theme_manager.get_button_colors(variant)
                self.assertIsInstance(colors, dict)
                # 按鈕顏色應該包含基本的 fg_color
                self.assertIn('fg_color', colors)


class TestCollapsibleModule(unittest.TestCase):
    """測試折疊組件模組"""
    
    def test_collapsible_imports(self):
        """測試折疊組件可正常匯入"""
        from ui.collapsible import CollapsibleSection
        self.assertIsNotNone(CollapsibleSection)
    
    def test_collapsible_class_methods(self):
        """測試 CollapsibleSection 類別具有必要的方法"""
        from ui.collapsible import CollapsibleSection
        
        required_methods = [
            'toggle', 'expand', 'collapse', 
            'is_expanded', 'get_content_frame', 'set_title'
        ]
        
        for method in required_methods:
            with self.subTest(method=method):
                self.assertTrue(
                    hasattr(CollapsibleSection, method),
                    f"CollapsibleSection should have method {method}"
                )


class TestUtilsModule(unittest.TestCase):
    """測試工具模組"""
    
    def test_utils_imports(self):
        """測試工具模組可正常匯入"""
        from app.utils import Tooltip, center_window, create_mount_directory, open_directory
        self.assertIsNotNone(Tooltip)
        self.assertIsNotNone(center_window)
    
    def test_create_mount_directory_function(self):
        """測試 create_mount_directory 函數簽名"""
        from app.utils import create_mount_directory
        import inspect
        sig = inspect.signature(create_mount_directory)
        params = list(sig.parameters.keys())
        self.assertIn('path', params)


class TestMainModernModule(unittest.TestCase):
    """測試主程式模組（不啟動 GUI）"""
    
    def test_main_modern_imports(self):
        """測試主程式模組可正常匯入"""
        # 只測試匯入，不創建應用程式實例
        import main_modern
        self.assertTrue(hasattr(main_modern, 'ModernApp'))
        self.assertTrue(hasattr(main_modern, 'main'))
    
    def test_modern_app_class_constants(self):
        """測試 ModernApp 類別常數定義"""
        from main_modern import ModernApp
        
        self.assertTrue(hasattr(ModernApp, 'APP_TITLE'))
        self.assertTrue(hasattr(ModernApp, 'APP_VERSION'))
        self.assertTrue(hasattr(ModernApp, 'WINDOW_SIZE'))
        self.assertTrue(hasattr(ModernApp, 'MIN_SIZE'))
        self.assertTrue(hasattr(ModernApp, 'WIM_SLOT_COUNT'))
        
        # 檢查類型
        self.assertIsInstance(ModernApp.APP_TITLE, str)
        self.assertIsInstance(ModernApp.APP_VERSION, str)
        self.assertIsInstance(ModernApp.WINDOW_SIZE, str)
        self.assertIsInstance(ModernApp.MIN_SIZE, tuple)
        self.assertIsInstance(ModernApp.WIM_SLOT_COUNT, int)
        self.assertGreater(ModernApp.WIM_SLOT_COUNT, 0)


class TestWIMPageModule(unittest.TestCase):
    """測試 WIM 頁面模組"""
    
    def test_wim_page_imports(self):
        """測試 WIM 頁面模組可正常匯入"""
        from ui.pages.wim_page import WIMSlot, WIMPage
        self.assertIsNotNone(WIMSlot)
        self.assertIsNotNone(WIMPage)
    
    def test_wim_slot_public_methods(self):
        """測試 WIMSlot 公開方法存在"""
        from ui.pages.wim_page import WIMSlot
        
        required_methods = [
            'get_config', 'set_config', 
            'get_mount_dir', 'is_mounted', 'check_status'
        ]
        
        for method in required_methods:
            with self.subTest(method=method):
                self.assertTrue(
                    hasattr(WIMSlot, method),
                    f"WIMSlot should have method {method}"
                )


class TestDriverPageModule(unittest.TestCase):
    """測試驅動程式頁面模組"""
    
    def test_driver_page_imports(self):
        """測試驅動程式頁面模組可正常匯入"""
        from ui.pages.driver_page import DriverPage
        self.assertIsNotNone(DriverPage)
    
    def test_driver_page_has_show_header_param(self):
        """測試 DriverPage 支援 show_header 參數"""
        from ui.pages.driver_page import DriverPage
        import inspect
        sig = inspect.signature(DriverPage.__init__)
        params = list(sig.parameters.keys())
        self.assertIn('show_header', params)


class TestComponentsModule(unittest.TestCase):
    """測試 UI 組件模組"""
    
    def test_components_imports(self):
        """測試組件模組可正常匯入"""
        from ui.components import (
            ModernButton, ModernCard, ModernEntry, 
            StatusBadge, IconButton, ModernTooltip
        )
        self.assertIsNotNone(ModernButton)
        self.assertIsNotNone(ModernCard)
        self.assertIsNotNone(StatusBadge)
    
    def test_modern_button_variants(self):
        """測試 ModernButton 支援所有變體"""
        from ui.components import ModernButton
        
        # 檢查類別存在 variant 相關邏輯
        import inspect
        source = inspect.getsource(ModernButton.__init__)
        variants = ['primary', 'secondary', 'success', 'warning', 'danger', 'outline', 'ghost']
        
        # 確認 variant 參數存在
        sig = inspect.signature(ModernButton.__init__)
        params = list(sig.parameters.keys())
        self.assertIn('variant', params)
    
    def test_status_badge_methods(self):
        """測試 StatusBadge 方法"""
        from ui.components import StatusBadge
        
        self.assertTrue(hasattr(StatusBadge, 'set_status'))


class TestLogPanelModule(unittest.TestCase):
    """測試日誌面板模組"""
    
    def test_log_panel_imports(self):
        """測試日誌面板模組可正常匯入"""
        from ui.log_panel import CollapsibleLogPanel
        self.assertIsNotNone(CollapsibleLogPanel)
    
    def test_log_panel_methods(self):
        """測試 CollapsibleLogPanel 具有必要方法"""
        from ui.log_panel import CollapsibleLogPanel
        
        required_methods = ['log', 'clear', 'get_content']
        
        for method in required_methods:
            with self.subTest(method=method):
                self.assertTrue(
                    hasattr(CollapsibleLogPanel, method),
                    f"CollapsibleLogPanel should have method {method}"
                )


class TestIntegration(unittest.TestCase):
    """整合測試"""
    
    def test_all_imports_work_together(self):
        """測試所有模組可以一起匯入"""
        try:
            from app.config import CONFIG_FILE, LOG_DIR, ConfigManager
            from app.utils import Tooltip
            from ui.theme import theme_manager, Colors, Fonts
            from ui.collapsible import CollapsibleSection
            from ui.components import ModernButton, ModernCard, StatusBadge
            from ui.log_panel import CollapsibleLogPanel
            from ui.pages.wim_page import WIMSlot
            from ui.pages.driver_page import DriverPage
            import main_modern
            
            success = True
        except ImportError as e:
            success = False
            self.fail(f"Import failed: {e}")
        
        self.assertTrue(success)
    
    def test_config_dir_constants_consistency(self):
        """測試設定檔路徑常數一致性"""
        from app.config import SCRIPT_DIR, CONFIG_FILE, LOG_DIR, OUTPUT_DIR
        
        # 所有路徑都應該以 SCRIPT_DIR 為基礎
        self.assertTrue(CONFIG_FILE.startswith(SCRIPT_DIR) or 
                       os.path.dirname(CONFIG_FILE) == SCRIPT_DIR.rstrip(os.sep))
        self.assertTrue(LOG_DIR.startswith(SCRIPT_DIR))
        self.assertTrue(OUTPUT_DIR.startswith(SCRIPT_DIR))


def run_tests():
    """執行所有測試"""
    # 建立測試套件
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # 添加所有測試類別
    test_classes = [
        TestConfigModule,
        TestThemeModule,
        TestCollapsibleModule,
        TestUtilsModule,
        TestMainModernModule,
        TestWIMPageModule,
        TestDriverPageModule,
        TestComponentsModule,
        TestLogPanelModule,
        TestIntegration,
    ]
    
    for test_class in test_classes:
        tests = loader.loadTestsFromTestCase(test_class)
        suite.addTests(tests)
    
    # 執行測試
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result


if __name__ == '__main__':
    result = run_tests()
    
    # 輸出摘要
    print("\n" + "=" * 60)
    print("測試摘要")
    print("=" * 60)
    print(f"執行測試數: {result.testsRun}")
    print(f"成功: {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"失敗: {len(result.failures)}")
    print(f"錯誤: {len(result.errors)}")
    
    if result.failures:
        print("\n失敗的測試:")
        for test, traceback in result.failures:
            print(f"  - {test}")
    
    if result.errors:
        print("\n錯誤的測試:")
        for test, traceback in result.errors:
            print(f"  - {test}")
    
    # 返回退出碼
    sys.exit(0 if result.wasSuccessful() else 1)
