# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WIM 掛載功能詳細測試單元

測試範圍：
1. WIMManager 核心功能
2. WIMSlot UI 組件
3. 設定管理功能
4. 狀態檢查功能
5. 錯誤處理機制
"""

import unittest
import sys
import os
import tempfile
import shutil
from unittest.mock import Mock, patch, MagicMock

# 確保專案根目錄在路徑中
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)
os.chdir(PROJECT_ROOT)


# ========== WIMManager 核心功能測試 ==========

class TestWIMManagerBasics(unittest.TestCase):
    """測試 WIMManager 基礎功能"""
    
    def test_wim_manager_import(self):
        """測試 WIMManager 可正常匯入"""
        from app.wim_manager import WIMManager
        self.assertIsNotNone(WIMManager)
    
    def test_norm_path(self):
        """測試路徑正規化"""
        from app.wim_manager import WIMManager
        
        # 測試正常路徑
        result = WIMManager._norm_path("C:/test/path")
        self.assertIsInstance(result, str)
        self.assertIn("test", result)
        
        # 測試空路徑
        result = WIMManager._norm_path("")
        self.assertEqual(result, ".")
        
        # 測試帶斜線的路徑
        result = WIMManager._norm_path("C:\\test\\path\\")
        self.assertIn("test", result)
    
    def test_is_admin_returns_bool(self):
        """測試 is_admin 返回布林值"""
        from app.wim_manager import WIMManager
        
        result = WIMManager.is_admin()
        self.assertIsInstance(result, bool)


class TestWIMManagerParseWimInfo(unittest.TestCase):
    """測試 WIMManager 解析 WIM 資訊功能"""
    
    def test_parse_wiminfo_basic(self):
        """測試基本 WIM 資訊解析"""
        from app.wim_manager import WIMManager
        
        sample_output = """
Deployment Image Servicing and Management tool
Version: 10.0.19041.1

Details for image : C:\\test.wim

Index : 1
Name : Windows 10 Pro
Description : Windows 10 Pro
Size : 15,432,123,456 bytes

Index : 2
Name : Windows 10 Home
Description : Windows 10 Home Edition
Size : 14,123,456,789 bytes
"""
        
        result = WIMManager._parse_wiminfo(sample_output)
        
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 2)
        
        # 檢查第一個映像
        self.assertEqual(result[0]["Index"], 1)
        self.assertEqual(result[0]["Name"], "Windows 10 Pro")
        self.assertEqual(result[0]["Description"], "Windows 10 Pro")
        
        # 檢查第二個映像
        self.assertEqual(result[1]["Index"], 2)
        self.assertEqual(result[1]["Name"], "Windows 10 Home")
    
    def test_parse_wiminfo_empty(self):
        """測試空輸出解析"""
        from app.wim_manager import WIMManager
        
        result = WIMManager._parse_wiminfo("")
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 0)
    
    def test_parse_wiminfo_single_image(self):
        """測試單一映像解析"""
        from app.wim_manager import WIMManager
        
        sample_output = """
Index : 1
Name : Test Image
Description : Test Description
"""
        
        result = WIMManager._parse_wiminfo(sample_output)
        
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["Index"], 1)
        self.assertEqual(result[0]["Name"], "Test Image")


class TestWIMManagerParseMountedInfo(unittest.TestCase):
    """測試 WIMManager 解析掛載資訊功能"""
    
    def test_parse_mounted_info_basic(self):
        """測試基本掛載資訊解析"""
        from app.wim_manager import WIMManager
        
        sample_output = """
Deployment Image Servicing and Management tool
Version: 10.0.19041.1

Mounted Images:

Mount Dir : C:\\MountDir1
Image File : C:\\test1.wim
Image Index : 1
Mounted Read/Write : Yes
Status : Ok

Mount Dir : C:\\MountDir2
Image File : C:\\test2.wim
Image Index : 2
Mounted Read Only : Yes
Status : Ok
"""
        
        result = WIMManager._parse_mounted_info(sample_output)
        
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 2)
        
        # 檢查第一個掛載
        self.assertEqual(result[0]["MountDir"], "C:\\MountDir1")
        self.assertEqual(result[0]["ImageFile"], "C:\\test1.wim")
    
    def test_parse_mounted_info_empty(self):
        """測試無掛載映像解析"""
        from app.wim_manager import WIMManager
        
        sample_output = """
Deployment Image Servicing and Management tool
No mounted images.
"""
        
        result = WIMManager._parse_mounted_info(sample_output)
        self.assertIsInstance(result, list)


class TestWIMManagerDISMExecution(unittest.TestCase):
    """測試 WIMManager DISM 執行功能"""
    
    @patch('subprocess.run')
    def test_run_dism_success(self, mock_run):
        """測試 DISM 執行成功"""
        from app.wim_manager import WIMManager
        
        mock_run.return_value = MagicMock(
            returncode=0,
            stdout="Success output",
            stderr=""
        )
        
        rc, out, err = WIMManager._run_dism(["/test"])
        
        self.assertEqual(rc, 0)
        self.assertEqual(out, "Success output")
        self.assertEqual(err, "")
    
    @patch('subprocess.run')
    def test_run_dism_failure(self, mock_run):
        """測試 DISM 執行失敗"""
        from app.wim_manager import WIMManager
        
        mock_run.return_value = MagicMock(
            returncode=1,
            stdout="",
            stderr="Error occurred"
        )
        
        rc, out, err = WIMManager._run_dism(["/test"])
        
        self.assertEqual(rc, 1)
        self.assertIn("Error", err)
    
    @patch('subprocess.run')
    def test_run_dism_file_not_found(self, mock_run):
        """測試 DISM 找不到情況"""
        from app.wim_manager import WIMManager
        
        mock_run.side_effect = FileNotFoundError("dism not found")
        
        rc, out, err = WIMManager._run_dism(["/test"])
        
        self.assertEqual(rc, 9001)
        self.assertIn("找不到 DISM", err)


class TestWIMManagerMountStatus(unittest.TestCase):
    """測試 WIMManager 掛載狀態功能"""
    
    def test_get_mount_status_for_path_nonexistent(self):
        """測試不存在路徑的狀態"""
        from app.wim_manager import WIMManager
        
        # 使用不存在的路徑
        status, details = WIMManager.get_mount_status_for_path("C:\\NonExistentPath12345")
        
        # 在沒有管理員權限時可能返回 error，否則返回 not_mounted
        self.assertIn(status, ["not_mounted", "error"])
        if status == "not_mounted":
            self.assertIn("不存在", details)
    
    @patch.object(__import__('app.wim_manager', fromlist=['WIMManager']).WIMManager, 'is_path_mounted')
    def test_get_mount_status_mounted(self, mock_is_mounted):
        """測試已掛載路徑的狀態"""
        from app.wim_manager import WIMManager
        
        mock_is_mounted.return_value = (True, {"Status": "Ok", "ReadWrite": "Read/Write"}, "")
        
        status, details = WIMManager.get_mount_status_for_path("C:\\Test")
        
        self.assertEqual(status, "mounted")
        self.assertIn("已掛載", details)
    
    @patch.object(__import__('app.wim_manager', fromlist=['WIMManager']).WIMManager, 'is_path_mounted')
    def test_get_mount_status_needs_remount(self, mock_is_mounted):
        """測試需要重新掛載的狀態"""
        from app.wim_manager import WIMManager
        
        mock_is_mounted.return_value = (True, {"Status": "Needs Remount", "ReadWrite": ""}, "")
        
        status, details = WIMManager.get_mount_status_for_path("C:\\Test")
        
        self.assertEqual(status, "needs_remount")


# ========== WIMSlot UI 組件測試 ==========

class TestWIMSlotBasics(unittest.TestCase):
    """測試 WIMSlot 基礎功能"""
    
    @classmethod
    def setUpClass(cls):
        """建立測試用的 root 視窗"""
        import customtkinter as ctk
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        """清理測試視窗"""
        cls.root.destroy()
    
    def test_wim_slot_creation(self):
        """測試建立 WIMSlot"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        
        slot = WIMSlot(
            self.root,
            slot_number=1,
            on_log=mock_log
        )
        
        self.assertIsNotNone(slot)
        self.assertEqual(slot.slot_number, 1)
        
        slot.destroy()
    
    def test_wim_slot_variables_initialization(self):
        """測試 WIMSlot 變數初始化"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        
        slot = WIMSlot(
            self.root,
            slot_number=2,
            on_log=mock_log
        )
        
        # 測試預設值
        self.assertEqual(slot.var_wim_path.get(), "")
        self.assertEqual(slot.var_mount_dir.get(), "")
        self.assertEqual(slot.var_index.get(), "")
        self.assertTrue(slot.var_readonly.get())  # 預設唯讀
        self.assertFalse(slot.var_commit.get())   # 預設丟棄變更
        self.assertEqual(slot.var_status.get(), "未檢查")
        
        slot.destroy()


class TestWIMSlotConfig(unittest.TestCase):
    """測試 WIMSlot 設定管理"""
    
    @classmethod
    def setUpClass(cls):
        import customtkinter as ctk
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_get_config(self):
        """測試取得設定"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        # 設定一些值
        slot.var_wim_path.set("C:\\test.wim")
        slot.var_mount_dir.set("C:\\Mount")
        slot.var_index.set("1")
        slot.var_readonly.set(False)
        slot.var_commit.set(True)
        
        config = slot.get_config()
        
        self.assertIsInstance(config, dict)
        self.assertEqual(config["wim_file"], "C:\\test.wim")
        self.assertEqual(config["mount_dir"], "C:\\Mount")
        self.assertEqual(config["index"], "1")
        self.assertFalse(config["readonly"])
        self.assertTrue(config["commit"])
        
        slot.destroy()
    
    def test_set_config(self):
        """測試設定配置"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        config = {
            "wim_file": "D:\\image.wim",
            "mount_dir": "D:\\MountPoint",
            "index": "3",
            "readonly": True,
            "commit": False
        }
        
        slot.set_config(config)
        
        self.assertEqual(slot.var_wim_path.get(), "D:\\image.wim")
        self.assertEqual(slot.var_mount_dir.get(), "D:\\MountPoint")
        self.assertEqual(slot.var_index.get(), "3")
        self.assertTrue(slot.var_readonly.get())
        self.assertFalse(slot.var_commit.get())
        
        slot.destroy()
    
    def test_set_config_partial(self):
        """測試部分設定配置"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        # 先設定一些初始值
        slot.var_wim_path.set("original.wim")
        
        # 部分設定
        config = {
            "mount_dir": "E:\\NewMount",
        }
        
        slot.set_config(config)
        
        # wim_file 應該保持原值
        self.assertEqual(slot.var_wim_path.get(), "original.wim")
        # mount_dir 應該更新
        self.assertEqual(slot.var_mount_dir.get(), "E:\\NewMount")
        
        slot.destroy()


class TestWIMSlotPublicMethods(unittest.TestCase):
    """測試 WIMSlot 公開方法"""
    
    @classmethod
    def setUpClass(cls):
        import customtkinter as ctk
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_get_mount_dir(self):
        """測試取得掛載目錄"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        slot.var_mount_dir.set("  C:\\TestMount  ")  # 帶空格
        
        result = slot.get_mount_dir()
        
        self.assertEqual(result, "C:\\TestMount")  # 應該去除空格
        
        slot.destroy()
    
    def test_is_mounted_default(self):
        """測試預設掛載狀態"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        # 預設應該是未掛載
        self.assertFalse(slot.is_mounted())
        
        slot.destroy()
    
    def test_is_mounted_after_status_change(self):
        """測試狀態變更後的掛載狀態"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        # 模擬狀態變更
        slot.var_status.set("已掛載")
        self.assertTrue(slot.is_mounted())
        
        slot.var_status.set("未掛載")
        self.assertFalse(slot.is_mounted())
        
        slot.destroy()


class TestWIMSlotCallbacks(unittest.TestCase):
    """測試 WIMSlot 回調機制"""
    
    @classmethod
    def setUpClass(cls):
        import customtkinter as ctk
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_on_log_callback(self):
        """測試日誌回調"""
        from ui.pages.wim_page import WIMSlot
        
        log_messages = []
        
        def mock_log(msg):
            log_messages.append(msg)
        
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        # 觸發日誌
        slot._on_log("測試訊息")
        
        self.assertEqual(len(log_messages), 1)
        self.assertEqual(log_messages[0], "測試訊息")
        
        slot.destroy()
    
    def test_on_status_change_callback(self):
        """測試狀態變更回調"""
        from ui.pages.wim_page import WIMSlot
        
        status_changes = []
        
        def mock_status_change(slot_number, is_mounted):
            status_changes.append((slot_number, is_mounted))
        
        mock_log = Mock()
        slot = WIMSlot(
            self.root,
            slot_number=1,
            on_log=mock_log,
            on_status_change=mock_status_change
        )
        
        # 模擬狀態更新
        slot._update_status("已掛載", "success")
        
        self.assertEqual(len(status_changes), 1)
        self.assertEqual(status_changes[0], (1, True))
        
        slot.destroy()
    
    def test_get_other_slot_info_callback(self):
        """測試取得其他槽位資訊回調"""
        from ui.pages.wim_page import WIMSlot
        
        def mock_get_other():
            return {"wim_file": "other.wim", "mount_dir": "C:\\Other"}
        
        mock_log = Mock()
        slot = WIMSlot(
            self.root,
            slot_number=1,
            on_log=mock_log,
            get_other_slot_info=mock_get_other
        )
        
        # 測試回調函數已正確設定
        self.assertIsNotNone(slot._get_other_slot_info)
        
        result = slot._get_other_slot_info()
        self.assertEqual(result["wim_file"], "other.wim")
        
        slot.destroy()


class TestWIMSlotFilePathComparison(unittest.TestCase):
    """測試 WIMSlot 檔案路徑比較功能"""
    
    @classmethod
    def setUpClass(cls):
        import customtkinter as ctk
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_is_same_file_identical(self):
        """測試相同路徑比較"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        result = slot._is_same_file("C:\\test.wim", "C:\\test.wim")
        self.assertTrue(result)
        
        slot.destroy()
    
    def test_is_same_file_different_case(self):
        """測試不同大小寫路徑比較"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        result = slot._is_same_file("C:\\TEST.wim", "c:\\test.wim")
        self.assertTrue(result)
        
        slot.destroy()
    
    def test_is_same_file_different_separators(self):
        """測試不同路徑分隔符號比較"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        result = slot._is_same_file("C:/test/file.wim", "C:\\test\\file.wim")
        self.assertTrue(result)
        
        slot.destroy()
    
    def test_is_same_file_different_files(self):
        """測試不同檔案比較"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        result = slot._is_same_file("C:\\test1.wim", "C:\\test2.wim")
        self.assertFalse(result)
        
        slot.destroy()
    
    def test_is_same_file_empty_path(self):
        """測試空路徑比較"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        result = slot._is_same_file("", "C:\\test.wim")
        self.assertFalse(result)
        
        result = slot._is_same_file("C:\\test.wim", "")
        self.assertFalse(result)
        
        slot.destroy()


class TestWIMSlotStatusUpdate(unittest.TestCase):
    """測試 WIMSlot 狀態更新功能"""
    
    @classmethod
    def setUpClass(cls):
        import customtkinter as ctk
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_update_status_success(self):
        """測試更新成功狀態"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        slot._update_status("已掛載", "success")
        
        self.assertEqual(slot.var_status.get(), "已掛載")
        # 按鈕狀態應該更新
        self.assertEqual(slot.btn_mount.cget("state"), "disabled")
        self.assertEqual(slot.btn_unmount.cget("state"), "normal")
        
        slot.destroy()
    
    def test_update_status_not_mounted(self):
        """測試更新未掛載狀態"""
        from ui.pages.wim_page import WIMSlot
        
        mock_log = Mock()
        slot = WIMSlot(self.root, slot_number=1, on_log=mock_log)
        
        slot._update_status("未掛載", "default")
        
        self.assertEqual(slot.var_status.get(), "未掛載")
        # 按鈕狀態應該更新
        self.assertEqual(slot.btn_mount.cget("state"), "normal")
        self.assertEqual(slot.btn_unmount.cget("state"), "disabled")
        
        slot.destroy()


# ========== WIMPage 整合測試 ==========

class TestWIMPageBasics(unittest.TestCase):
    """測試 WIMPage 基礎功能"""
    
    @classmethod
    def setUpClass(cls):
        import customtkinter as ctk
        cls.root = ctk.CTk()
        cls.root.withdraw()
    
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
    
    def test_wim_page_creation(self):
        """測試建立 WIMPage"""
        from ui.pages.wim_page import WIMPage
        
        mock_log = Mock()
        
        page = WIMPage(
            self.root,
            on_log=mock_log
        )
        
        self.assertIsNotNone(page)
        self.assertIsNotNone(page.slot1)
        self.assertIsNotNone(page.slot2)
        
        page.destroy()
    
    def test_wim_page_get_config(self):
        """測試 WIMPage 取得設定"""
        from ui.pages.wim_page import WIMPage
        
        mock_log = Mock()
        page = WIMPage(self.root, on_log=mock_log)
        
        config = page.get_config()
        
        self.assertIsInstance(config, dict)
        self.assertIn("WIM", config)
        self.assertIn("WIM2", config)
        
        page.destroy()
    
    def test_wim_page_set_config(self):
        """測試 WIMPage 設定配置"""
        from ui.pages.wim_page import WIMPage
        
        mock_log = Mock()
        page = WIMPage(self.root, on_log=mock_log)
        
        config = {
            "WIM": {
                "wim_file": "C:\\test1.wim",
                "mount_dir": "C:\\Mount1",
                "index": "1",
            },
            "WIM2": {
                "wim_file": "C:\\test2.wim",
                "mount_dir": "C:\\Mount2",
                "index": "2",
            }
        }
        
        page.set_config(config)
        
        # 驗證設定已套用
        self.assertEqual(page.slot1.var_wim_path.get(), "C:\\test1.wim")
        self.assertEqual(page.slot2.var_wim_path.get(), "C:\\test2.wim")
        
        page.destroy()
    
    def test_wim_page_get_mounted_dirs(self):
        """測試取得已掛載目錄"""
        from ui.pages.wim_page import WIMPage
        
        mock_log = Mock()
        page = WIMPage(self.root, on_log=mock_log)
        
        # 預設應該沒有掛載
        dirs = page.get_mounted_dirs()
        self.assertIsInstance(dirs, list)
        self.assertEqual(len(dirs), 0)
        
        page.destroy()


# ========== 測試執行器 ==========

def run_wim_tests():
    """執行所有 WIM 掛載相關測試"""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    test_classes = [
        # WIMManager 測試
        TestWIMManagerBasics,
        TestWIMManagerParseWimInfo,
        TestWIMManagerParseMountedInfo,
        TestWIMManagerDISMExecution,
        TestWIMManagerMountStatus,
        
        # WIMSlot 測試
        TestWIMSlotBasics,
        TestWIMSlotConfig,
        TestWIMSlotPublicMethods,
        TestWIMSlotCallbacks,
        TestWIMSlotFilePathComparison,
        TestWIMSlotStatusUpdate,
        
        # WIMPage 測試
        TestWIMPageBasics,
    ]
    
    for test_class in test_classes:
        tests = loader.loadTestsFromTestCase(test_class)
        suite.addTests(tests)
    
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result


if __name__ == '__main__':
    result = run_wim_tests()
    
    print("\n" + "=" * 70)
    print("               WIM 掛載功能測試摘要")
    print("=" * 70)
    print(f"  執行測試數:     {result.testsRun}")
    passed = result.testsRun - len(result.failures) - len(result.errors)
    print(f"  成功:           {passed} {'[PASS]' if passed == result.testsRun else ''}")
    print(f"  失敗:           {len(result.failures)} {'[FAIL]' if result.failures else ''}")
    print(f"  錯誤:           {len(result.errors)} {'[ERROR]' if result.errors else ''}")
    print("=" * 70)
    
    if result.failures:
        print("\n失敗的測試:")
        for test, traceback in result.failures:
            print(f"\n  ❌ {test}")
            lines = traceback.strip().split('\n')
            for line in lines[-5:]:
                print(f"     {line}")
    
    if result.errors:
        print("\n錯誤的測試:")
        for test, traceback in result.errors:
            print(f"\n  ❌ {test}")
            lines = traceback.strip().split('\n')
            for line in lines[-5:]:
                print(f"     {line}")
    
    sys.exit(0 if result.wasSuccessful() else 1)
