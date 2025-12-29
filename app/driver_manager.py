# -*- coding: utf-8 -*-
"""
Driver 管理模組 - 使用 DISM 進行驅動程式離線安裝/移除操作
"""

import os
import re
import subprocess


class DriverManager:
    """驅動程式管理類別 - 封裝所有驅動相關的 DISM 操作"""
    
    # ==================== 基礎工具方法 ====================
    
    @staticmethod
    def _norm_path(p: str) -> str:
        """正規化路徑"""
        try:
            return os.path.normpath(p)
        except Exception:
            return p

    @staticmethod
    def _run_dism(args: list[str]) -> tuple[int, str, str]:
        """
        執行 DISM 命令
        
        Returns:
            (return_code, stdout, stderr)
        """
        try:
            cp = subprocess.run(["dism", "/English", *args], capture_output=True, text=True)
            return cp.returncode, cp.stdout or "", cp.stderr or ""
        except FileNotFoundError as e:
            return 9001, "", f"找不到 DISM：{e}"
        except Exception as e:
            return 9002, "", str(e)

    # ==================== 驅動程式安裝 ====================

    @staticmethod
    def add_driver_to_offline_image(mount_dir: str, driver_path: str, recurse: bool = True, force_unsigned: bool = False) -> tuple[bool, str]:
        """
        離線安裝驅動程式到已掛載的映像
        
        Args:
            mount_dir: 映像掛載路徑
            driver_path: 驅動程式路徑（.inf 檔案或資料夾）
            recurse: 是否遞迴搜尋子資料夾（僅對資料夾有效）
            force_unsigned: 是否強制安裝未簽署的驅動
        
        Returns:
            (success, message)
        """
        m = DriverManager._norm_path(mount_dir)
        d = DriverManager._norm_path(driver_path)
        
        args = [
            "/Add-Driver",
            f"/Image:{m}",
            f"/Driver:{d}",
        ]
        
        # /Recurse 只能用於資料夾，不能用於單一 .inf 檔案
        is_inf_file = d.lower().endswith('.inf') and os.path.isfile(d)
        if recurse and not is_inf_file:
            args.append("/Recurse")
        
        if force_unsigned:
            args.append("/ForceUnsigned")
            
        rc, out, err = DriverManager._run_dism(args)
        if rc == 0:
            return True, "驅動程式安裝完成"
        return False, err or out

    # ==================== 驅動程式萃取 ====================

    @staticmethod  
    def export_drivers_from_offline_image(mount_dir: str, export_dir: str) -> tuple[bool, str]:
        """
        從已掛載的映像中萃取所有驅動程式
        
        Args:
            mount_dir: 映像掛載路徑
            export_dir: 萃取輸出目錄
        
        Returns:
            (success, message)
        """
        m = DriverManager._norm_path(mount_dir)
        e = DriverManager._norm_path(export_dir)
        
        # 確保匯出目錄存在
        os.makedirs(e, exist_ok=True)
        
        args = [
            "/Export-Driver",
            f"/Image:{m}",
            f"/Destination:{e}"
        ]
            
        rc, out, err = DriverManager._run_dism(args)
        if rc == 0:
            return True, "驅動程式萃取完成"
        return False, err or out

    @staticmethod
    def export_selected_drivers(mount_dir: str, export_dir: str, driver_names: list[str],
                                 callback=None) -> tuple[int, int, list[str]]:
        """
        從已掛載的映像中萃取選定的驅動程式
        
        Args:
            mount_dir: 映像掛載路徑
            export_dir: 萃取輸出目錄
            driver_names: 驅動程式名稱列表 (Published Name，如 oem0.inf)
            callback: 進度回調函數 (current, total, name, success, msg)
        
        Returns:
            (success_count, fail_count, errors)
        """
        m = DriverManager._norm_path(mount_dir)
        e = DriverManager._norm_path(export_dir)
        
        # 確保匯出目錄存在
        os.makedirs(e, exist_ok=True)
        
        success_count = 0
        fail_count = 0
        errors = []
        total = len(driver_names)
        
        for i, name in enumerate(driver_names, 1):
            # DISM /Export-Driver 支援 /Driver 參數來指定特定驅動
            args = [
                "/Export-Driver",
                f"/Image:{m}",
                f"/Destination:{e}",
                f"/Driver:{name}"
            ]
            
            rc, out, err = DriverManager._run_dism(args)
            
            if rc == 0:
                success_count += 1
                if callback:
                    callback(i, total, name, True, "")
            else:
                fail_count += 1
                error_msg = err or out
                errors.append(f"{name}: {error_msg}")
                if callback:
                    callback(i, total, name, False, error_msg)
        
        return success_count, fail_count, errors

    # ==================== 驅動程式移除 ====================

    @staticmethod
    def remove_driver_from_offline_image(mount_dir: str, driver_name: str) -> tuple[bool, str]:
        """
        從已掛載的映像中移除指定的驅動程式
        
        Args:
            mount_dir: 映像掛載路徑
            driver_name: 驅動程式名稱 (Published Name，如 oem0.inf)
        
        Returns:
            (success, message)
        """
        m = DriverManager._norm_path(mount_dir)
        
        args = [
            "/Remove-Driver",
            f"/Image:{m}",
            f"/Driver:{driver_name}"
        ]
        
        rc, out, err = DriverManager._run_dism(args)
        if rc == 0:
            return True, f"驅動程式 {driver_name} 移除成功"
        return False, err or out

    @staticmethod
    def remove_drivers_batch(mount_dir: str, driver_names: list[str], callback=None) -> tuple[int, int, list[str]]:
        """
        批量移除多個驅動程式
        
        Args:
            mount_dir: 映像掛載路徑
            driver_names: 驅動程式名稱列表
            callback: 進度回調函數 callback(current, total, driver_name, success, message)
        
        Returns:
            (success_count, fail_count, error_messages)
        """
        success_count = 0
        fail_count = 0
        errors = []
        total = len(driver_names)
        
        for i, driver_name in enumerate(driver_names):
            ok, msg = DriverManager.remove_driver_from_offline_image(mount_dir, driver_name)
            if ok:
                success_count += 1
            else:
                fail_count += 1
                errors.append(f"{driver_name}: {msg}")
            
            if callback:
                callback(i + 1, total, driver_name, ok, msg)
        
        return success_count, fail_count, errors

    # ==================== 驅動程式資訊查詢 ====================

    @staticmethod
    def get_driver_details(mount_dir: str, driver_name: str) -> tuple[bool, dict, str]:
        """
        取得單一驅動程式的詳細資訊
        
        Args:
            mount_dir: 映像掛載路徑
            driver_name: 驅動程式名稱 (Published Name)
        
        Returns:
            (success, driver_info_dict, error_message)
        """
        m = DriverManager._norm_path(mount_dir)
        
        args = [
            "/Get-DriverInfo",
            f"/Image:{m}",
            f"/Driver:{driver_name}"
        ]
        
        rc, out, err = DriverManager._run_dism(args)
        if rc != 0:
            return False, {}, err or out
        
        # 解析輸出
        info = {"PublishedName": driver_name}
        for line in out.splitlines():
            line = line.strip()
            if ":" in line:
                key, _, value = line.partition(":")
                key = key.strip().replace(" ", "")
                value = value.strip()
                if key and value:
                    info[key] = value
        
        return True, info, ""

    @staticmethod
    def get_driver_info_from_path(driver_path: str) -> tuple[bool, list[dict], str]:
        """
        取得指定路徑中的驅動程式資訊（掃描 .inf 檔案）
        
        Args:
            driver_path: 驅動程式路徑（檔案或資料夾）
        
        Returns:
            (success, drivers_list, message)
        """
        if not os.path.exists(driver_path):
            return False, [], "路徑不存在"
            
        drivers = []
        try:
            if os.path.isfile(driver_path) and driver_path.lower().endswith('.inf'):
                # 單一 .inf 檔案
                driver_info = {"path": driver_path, "name": os.path.basename(driver_path)}
                drivers.append(driver_info)
            elif os.path.isdir(driver_path):
                # 資料夾，搜尋所有 .inf 檔案
                for root, dirs, files in os.walk(driver_path):
                    for file in files:
                        if file.lower().endswith('.inf'):
                            full_path = os.path.join(root, file)
                            driver_info = {"path": full_path, "name": file, "folder": root}
                            drivers.append(driver_info)
            
            return True, drivers, f"找到 {len(drivers)} 個驅動程式檔案"
        except Exception as e:
            return False, [], f"掃描驅動程式時發生錯誤: {str(e)}"

    @staticmethod
    def get_drivers_in_offline_image(mount_dir: str) -> tuple[bool, list[dict], str]:
        """
        列出已安裝在離線映像中的驅動程式
        
        Args:
            mount_dir: 映像掛載路徑
        
        Returns:
            (success, drivers_list, error_message)
        """
        m = DriverManager._norm_path(mount_dir)
        args = ["/Get-Drivers", f"/Image:{m}"]
        
        rc, out, err = DriverManager._run_dism(args)
        if rc != 0:
            return False, [], err or out
            
        drivers = DriverManager._parse_drivers(out)
        return True, drivers, ""

    @staticmethod
    def _parse_drivers(text: str) -> list[dict]:
        """
        解析 DISM 驅動程式輸出
        
        Args:
            text: DISM /Get-Drivers 的輸出文字
        
        Returns:
            驅動程式資訊列表
        """
        drivers: list[dict] = []
        cur: dict | None = None
        
        # 欄位對應表：DISM 輸出欄位名稱 -> dict key
        field_map = {
            "Published Name": "PublishedName",
            "Original File Name": "OriginalFileName",
            "Class Name": "ClassName",
            "Provider Name": "Provider",
            "Date": "Date",
            "Version": "Version",
        }
        
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
                
            # 檢測新驅動程式開始
            if line.startswith("Published Name"):
                if cur:
                    drivers.append(cur)
                cur = {"PublishedName": "", "OriginalFileName": "", "ClassName": "", "Provider": "", "Date": "", "Version": ""}
                match = re.search(r"Published Name\s*:\s*(.*)", line, re.IGNORECASE)
                if match:
                    cur["PublishedName"] = match.group(1).strip()
                continue
                
            if cur is not None:
                # 嘗試匹配各欄位
                for dism_field, dict_key in field_map.items():
                    if dict_key == "PublishedName":  # 已處理
                        continue
                    pattern = f"{re.escape(dism_field)}\\s*:\\s*(.*)"
                    match = re.search(pattern, line, re.IGNORECASE)
                    if match:
                        cur[dict_key] = match.group(1).strip()
                        break
                        
        if cur:
            drivers.append(cur)
            
        return drivers
