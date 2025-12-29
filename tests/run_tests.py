# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試執行器
執行所有單元測試和功能測試
"""

import unittest
import sys
import os
from datetime import datetime

# 確保專案根目錄在路徑中
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)
os.chdir(PROJECT_ROOT)


def run_all_tests(verbose: int = 2) -> dict:
    """
    執行所有測試
    
    Args:
        verbose: 詳細程度 (0=靜默, 1=簡短, 2=詳細)
    
    Returns:
        測試結果摘要字典
    """
    # 匯入測試模組
    from tests import test_modules, test_functional
    
    # 建立測試套件
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # 添加模組測試
    module_tests = [
        test_modules.TestConfigModule,
        test_modules.TestThemeModule,
        test_modules.TestCollapsibleModule,
        test_modules.TestUtilsModule,
        test_modules.TestMainModernModule,
        test_modules.TestWIMPageModule,
        test_modules.TestDriverPageModule,
        test_modules.TestComponentsModule,
        test_modules.TestLogPanelModule,
        test_modules.TestIntegration,
    ]
    
    # 添加功能測試
    functional_tests = [
        test_functional.TestCollapsibleSectionFunctionality,
        test_functional.TestModernCardFunctionality,
        test_functional.TestModernButtonFunctionality,
        test_functional.TestStatusBadgeFunctionality,
        test_functional.TestCollapsibleLogPanelFunctionality,
        test_functional.TestConfigManagerFunctionality,
        test_functional.TestThemeManagerFunctionality,
    ]
    
    all_test_classes = module_tests + functional_tests
    
    for test_class in all_test_classes:
        tests = loader.loadTestsFromTestCase(test_class)
        suite.addTests(tests)
    
    # 執行測試
    runner = unittest.TextTestRunner(verbosity=verbose)
    result = runner.run(suite)
    
    # 建立摘要
    summary = {
        'total': result.testsRun,
        'passed': result.testsRun - len(result.failures) - len(result.errors),
        'failures': len(result.failures),
        'errors': len(result.errors),
        'success': result.wasSuccessful(),
        'failure_details': result.failures,
        'error_details': result.errors,
    }
    
    return summary


def print_summary(summary: dict):
    """印出測試摘要"""
    print("\n" + "=" * 70)
    print("                          測試摘要報告")
    print("=" * 70)
    print(f"執行時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("-" * 70)
    print(f"  總測試數:     {summary['total']}")
    print(f"  成功:         {summary['passed']} ✅")
    print(f"  失敗:         {summary['failures']} {'❌' if summary['failures'] > 0 else ''}")
    print(f"  錯誤:         {summary['errors']} {'❌' if summary['errors'] > 0 else ''}")
    print("-" * 70)
    
    if summary['success']:
        print("  結果: 全部通過 ✅")
    else:
        print("  結果: 有測試未通過 ❌")
        
        if summary['failure_details']:
            print("\n失敗的測試:")
            for test, traceback in summary['failure_details']:
                print(f"\n  ❌ {test}")
                # 只印出關鍵錯誤訊息
                lines = traceback.strip().split('\n')
                for line in lines[-3:]:
                    print(f"     {line}")
        
        if summary['error_details']:
            print("\n錯誤的測試:")
            for test, traceback in summary['error_details']:
                print(f"\n  ❌ {test}")
                lines = traceback.strip().split('\n')
                for line in lines[-3:]:
                    print(f"     {line}")
    
    print("=" * 70)


def main():
    """主程式"""
    import argparse
    
    parser = argparse.ArgumentParser(description='執行單元測試套件')
    parser.add_argument('-v', '--verbose', type=int, default=2,
                       choices=[0, 1, 2], help='詳細程度 (0=靜默, 1=簡短, 2=詳細)')
    parser.add_argument('-q', '--quiet', action='store_true',
                       help='靜默模式 (等同於 -v 0)')
    parser.add_argument('--module-only', action='store_true',
                       help='只執行模組測試')
    parser.add_argument('--functional-only', action='store_true',
                       help='只執行功能測試')
    
    args = parser.parse_args()
    
    verbose = 0 if args.quiet else args.verbose
    
    print("=" * 70)
    print("          WIM/Driver 管理工具 - 單元測試套件")
    print("=" * 70)
    print()
    
    if args.module_only:
        print("執行模組測試...")
        from tests.test_modules import run_tests
        result = run_tests()
        success = result.wasSuccessful()
    elif args.functional_only:
        print("執行功能測試...")
        from tests.test_functional import run_functional_tests
        result = run_functional_tests()
        success = result.wasSuccessful()
    else:
        print("執行所有測試...")
        summary = run_all_tests(verbose)
        print_summary(summary)
        success = summary['success']
    
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
