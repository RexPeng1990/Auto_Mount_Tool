# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
功能測試套件 - 測試實際功能行為

這套測試會建立實際的 GUI 組件來測試功能
需要在有顯示環境的情況下執行
"""

import unittest
import sys
import os

# 確保專案根目錄在路徑中
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)
os.chdir(PROJECT_ROOT)

import tkinter as tk
import customtkinter as ctk


class TestCollapsibleSectionFunctionality(unittest.TestCase):
    """測試 CollapsibleSection 實際功能"""
    
    @classmethod
    def setUpClass(cls):
        """建立測試用的 root 視窗"""
        cls.root = ctk.CTk()
        cls.root.withdraw()  # 隱藏視窗
    
    @classmethod
    def tearDownClass(cls):
        """清理測試視窗"""
        cls.root.destroy()
    
    def test_collapsible_creation(self):
        """測試建立 CollapsibleSection"""
        from ui.collapsible import CollapsibleSection
        
        section = CollapsibleSection(
            self.root,
            title="測試區塊",
            icon="🔧",
            default_expanded=True
        )
        
        self.assertIsNotNone(section)
        self.assertTrue(section.is_expanded())
        section.destroy()
    
    def test_collapsible_toggle(self):
        """測試展開/收合切換"""
        from ui.collapsible import CollapsibleSection
        
        section = CollapsibleSection(
            self.root,
            title="測試區塊",
            default_expanded=True
        )
        
        self.assertTrue(section.is_expanded())
        
        section.toggle()
        self.assertFalse(section.is_expanded())
        
        section.toggle()
        self.assertTrue(section.is_expanded())
        
        section.destroy()
    
    def test_collapsible_expand_collapse(self):
        """測試 expand/collapse 方法"""
        from ui.collapsible import CollapsibleSection
        
        section = CollapsibleSection(
            self.root,
            title="測試區塊",
            default_expanded=False
        )
        
        self.assertFalse(section.is_expanded())
        
        section.expand()
        self.assertTrue(section.is_expanded())
        
        section.collapse()
        self.assertFalse(section.is_expanded())
        
        # 再次 collapse 應該無效果
        section.collapse()
        self.assertFalse(section.is_expanded())
        
        section.destroy()
    
    def test_collapsible_content_frame(self):
        """測試內容框架可用"""
        from ui.collapsible import CollapsibleSection
        
        section = CollapsibleSection(
            self.root,
            title="測試區塊",
            default_expanded=True
        )
        
        content = section.get_content_frame()
        self.assertIsNotNone(content)
        self.assertIsInstance(content, ctk.CTkFrame)
        
        # 測試可以在內容框架中添加元素
        label = ctk.CTkLabel(content, text="測試標籤")
        label.pack()
        
        section.destroy()
    
    def test_collapsible_set_title(self):
        """測試設定標題"""
        from ui.collapsible import CollapsibleSection
        
        section = CollapsibleSection(
            self.root,
            title="原始標題",
            icon="🔧"
        )
        
        section.set_title("新標題")
        # 由於標題包含圖示，檢查標題標籤的文字
        self.assertIn("新標題", section.title_label.cget("text"))
        
        section.destroy()
    
    def test_collapsible_callback(self):
        """測試 on_toggle 回調"""
        from ui.collapsible import CollapsibleSection
        
        callback_results = []
        
        def on_toggle(expanded):
            callback_results.append(expanded)
        
        section = CollapsibleSection(
            self.root,
            title="測試區塊",
            default_expanded=True,
            on_toggle=on_toggle
        )
        
        section.toggle()  # 收合
        section.toggle()  # 展開
        
        self.assertEqual(callback_results, [False, True])
        
        section.destroy()


class TestModernCardFunctionality(unittest.TestCase):
    """測試 ModernCard 實際功能"""
    
    @classmethod
    def setUpClass(cls):
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_modern_card_creation(self):
        """測試建立 ModernCard"""
        from ui.components import ModernCard
        
        card = ModernCard(self.root, padding=16)
        self.assertIsNotNone(card)
        
        content = card.get_content_frame()
        self.assertIsNotNone(content)
        
        card.destroy()
    
    def test_modern_card_content_frame(self):
        """測試 ModernCard 內容框架"""
        from ui.components import ModernCard
        
        card = ModernCard(self.root, padding=20)
        content = card.get_content_frame()
        
        # 測試可以添加子元素
        label = ctk.CTkLabel(content, text="測試")
        label.pack()
        
        card.destroy()


class TestModernButtonFunctionality(unittest.TestCase):
    """測試 ModernButton 實際功能"""
    
    @classmethod
    def setUpClass(cls):
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_button_creation_all_variants(self):
        """測試建立所有變體的按鈕"""
        from ui.components import ModernButton
        
        variants = ['primary', 'secondary', 'success', 'warning', 'danger', 'outline', 'ghost']
        
        for variant in variants:
            with self.subTest(variant=variant):
                btn = ModernButton(
                    self.root,
                    text=f"{variant} 按鈕",
                    variant=variant
                )
                self.assertIsNotNone(btn)
                btn.destroy()
    
    def test_button_sizes(self):
        """測試按鈕尺寸"""
        from ui.components import ModernButton
        
        sizes = ['sm', 'md', 'lg']
        
        for size in sizes:
            with self.subTest(size=size):
                btn = ModernButton(
                    self.root,
                    text="測試",
                    size=size
                )
                self.assertIsNotNone(btn)
                btn.destroy()
    
    def test_button_loading_state(self):
        """測試按鈕載入狀態"""
        from ui.components import ModernButton
        
        btn = ModernButton(self.root, text="測試按鈕")
        
        original_text = btn.cget("text")
        
        btn.set_loading(True)
        self.assertIn("處理中", btn.cget("text"))
        self.assertEqual(btn.cget("state"), "disabled")
        
        btn.set_loading(False)
        self.assertEqual(btn.cget("text"), original_text)
        self.assertEqual(btn.cget("state"), "normal")
        
        btn.destroy()
    
    def test_button_with_icon(self):
        """測試帶圖示的按鈕"""
        from ui.components import ModernButton
        
        btn = ModernButton(
            self.root,
            text="儲存",
            icon="💾"
        )
        
        self.assertIn("💾", btn.cget("text"))
        self.assertIn("儲存", btn.cget("text"))
        
        btn.destroy()


class TestStatusBadgeFunctionality(unittest.TestCase):
    """測試 StatusBadge 實際功能"""
    
    @classmethod
    def setUpClass(cls):
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_status_badge_creation(self):
        """測試建立 StatusBadge"""
        from ui.components import StatusBadge
        
        badge = StatusBadge(self.root, text="測試", status="default")
        self.assertIsNotNone(badge)
        badge.destroy()
    
    def test_status_badge_set_status(self):
        """測試設定狀態"""
        from ui.components import StatusBadge
        
        badge = StatusBadge(self.root, text="初始", status="default")
        
        statuses = ['success', 'warning', 'danger', 'info', 'default']
        
        for status in statuses:
            with self.subTest(status=status):
                badge.set_status(f"{status} 狀態", status)
                # 不應該拋出異常
        
        badge.destroy()


class TestCollapsibleLogPanelFunctionality(unittest.TestCase):
    """測試 CollapsibleLogPanel 實際功能"""
    
    @classmethod
    def setUpClass(cls):
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_log_panel_creation(self):
        """測試建立日誌面板"""
        from ui.log_panel import CollapsibleLogPanel
        
        panel = CollapsibleLogPanel(
            self.root,
            title="系統日誌",
            default_expanded=True
        )
        self.assertIsNotNone(panel)
        panel.destroy()
    
    def test_log_panel_log_message(self):
        """測試寫入日誌訊息"""
        from ui.log_panel import CollapsibleLogPanel
        
        panel = CollapsibleLogPanel(
            self.root,
            title="系統日誌",
            default_expanded=True
        )
        
        # 寫入訊息
        panel.log("測試訊息 1")
        panel.log("測試訊息 2", level="INFO")
        panel.log("警告訊息", level="WARNING")
        panel.log("錯誤訊息", level="ERROR")
        
        # 取得內容
        content = panel.get_content()
        self.assertIn("測試訊息 1", content)
        self.assertIn("測試訊息 2", content)
        
        panel.destroy()
    
    def test_log_panel_clear(self):
        """測試清除日誌"""
        from ui.log_panel import CollapsibleLogPanel
        
        panel = CollapsibleLogPanel(
            self.root,
            title="系統日誌",
            default_expanded=True
        )
        
        panel.log("測試訊息")
        panel.clear()
        
        content = panel.get_content()
        # 清除後內容應該為空或只有空白
        self.assertTrue(len(content.strip()) == 0 or content.strip() == "")
        
        panel.destroy()


class TestConfigManagerFunctionality(unittest.TestCase):
    """測試 ConfigManager 實際功能"""
    
    def test_config_set_and_get(self):
        """測試設定和取得設定值"""
        from app.config import ConfigManager
        import tempfile
        
        # 使用臨時檔案
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ini', delete=False) as f:
            temp_path = f.name
        
        try:
            cm = ConfigManager(config_file=temp_path)
            
            # 設定值
            cm.set('TEST', 'option1', 'value1')
            cm.set('TEST', 'option2', 'value2')
            cm.save()
            
            # 重新載入
            cm2 = ConfigManager(config_file=temp_path)
            
            self.assertEqual(cm2.get('TEST', 'option1'), 'value1')
            self.assertEqual(cm2.get('TEST', 'option2'), 'value2')
            
        finally:
            os.unlink(temp_path)


class TestThemeManagerFunctionality(unittest.TestCase):
    """測試 ThemeManager 實際功能"""
    
    def test_theme_switching(self):
        """測試主題切換"""
        from ui.theme import theme_manager, Colors
        
        # 測試設定深色主題
        theme_manager.set_theme('dark')
        self.assertEqual(theme_manager.colors.bg_primary, Colors.DARK.bg_primary)
        
        # 測試設定淺色主題
        theme_manager.set_theme('light')
        self.assertEqual(theme_manager.colors.bg_primary, Colors.LIGHT.bg_primary)
        
        # 恢復深色主題
        theme_manager.set_theme('dark')
    
    def test_button_colors_consistency(self):
        """測試按鈕顏色一致性"""
        from ui.theme import theme_manager
        
        colors1 = theme_manager.get_button_colors('primary')
        colors2 = theme_manager.get_button_colors('primary')
        
        # 相同變體應該返回相同的顏色配置
        self.assertEqual(colors1, colors2)


def run_functional_tests():
    """執行功能測試"""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    test_classes = [
        TestCollapsibleSectionFunctionality,
        TestModernCardFunctionality,
        TestModernButtonFunctionality,
        TestStatusBadgeFunctionality,
        TestCollapsibleLogPanelFunctionality,
        TestConfigManagerFunctionality,
        TestThemeManagerFunctionality,
    ]
    
    for test_class in test_classes:
        tests = loader.loadTestsFromTestCase(test_class)
        suite.addTests(tests)
    
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result


if __name__ == '__main__':
    result = run_functional_tests()
    
    print("\n" + "=" * 60)
    print("功能測試摘要")
    print("=" * 60)
    print(f"執行測試數: {result.testsRun}")
    print(f"成功: {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"失敗: {len(result.failures)}")
    print(f"錯誤: {len(result.errors)}")
    
    if result.failures:
        print("\n失敗的測試:")
        for test, traceback in result.failures:
            print(f"  - {test}")
            print(f"    {traceback}")
    
    if result.errors:
        print("\n錯誤的測試:")
        for test, traceback in result.errors:
            print(f"  - {test}")
            print(f"    {traceback}")
    
    sys.exit(0 if result.wasSuccessful() else 1)
