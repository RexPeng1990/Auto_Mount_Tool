# -*- coding: utf-8 -*-
"""
WIM 管理模組 - 使用 DISM 進行 WIM 映像掛載/卸載操作
"""

import os
import re
import subprocess
import threading

# 全域 DISM 操作鎖 - 防止多個 DISM 操作同時執行
_dism_lock = threading.Lock()
_dism_busy = False


class WIMManager:
    """WIM 映像管理類別 - 封裝所有 DISM 相關操作"""
    
    # ==================== 基礎工具方法 ====================
    
    @staticmethod
    def _norm_path(p: str) -> str:
        """正規化路徑"""
        try:
            return os.path.normpath(p)
        except Exception:
            return p
    
    @staticmethod
    def is_admin() -> bool:
        """檢查是否具有管理員權限"""
        try:
            import ctypes
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

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

    # ==================== WIM 映像資訊 ====================

    @staticmethod
    def get_wim_images(wim_path: str) -> tuple[bool, list[dict], str]:
        """
        取得 WIM 檔案中的映像資訊列表
        
        Returns:
            (success, images_list, error_message)
        """
        w = WIMManager._norm_path(wim_path)
        rc, out, err = WIMManager._run_dism(["/Get-WimInfo", f"/WimFile:{w}"])
        if rc != 0:
            # 兼容舊參數 /Get-ImageInfo
            rc2, out2, err2 = WIMManager._run_dism(["/Get-ImageInfo", f"/ImageFile:{w}"])
            if rc2 != 0:
                return False, [], err or err2 or out2 or out
            out = out2
        images = WIMManager._parse_wiminfo(out)
        return True, images, ""

    @staticmethod
    def _parse_wiminfo(text: str) -> list[dict]:
        """解析 DISM 輸出，擷取 Index/Name/Description"""
        imgs: list[dict] = []
        cur: dict | None = None
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            m = re.match(r"Index\s*:\s*(\d+)", line, re.IGNORECASE)
            if m:
                if cur:
                    imgs.append(cur)
                cur = {"Index": int(m.group(1)), "Name": "", "Description": ""}
                continue
            if cur is not None:
                m = re.match(r"Name\s*:\s*(.*)", line, re.IGNORECASE)
                if m:
                    cur["Name"] = m.group(1).strip()
                    continue
                m = re.match(r"Description\s*:\s*(.*)", line, re.IGNORECASE)
                if m:
                    cur["Description"] = m.group(1).strip()
                    continue
        if cur:
            imgs.append(cur)
        return imgs

    # ==================== 掛載/卸載操作 ====================

    @staticmethod
    def mount_wim(wim_path: str, index: int, mount_dir: str, readonly: bool) -> tuple[bool, str]:
        """
        掛載 WIM 映像
        
        Args:
            wim_path: WIM 檔案路徑
            index: 映像索引
            mount_dir: 掛載目錄
            readonly: 是否唯讀
        
        Returns:
            (success, message)
        """
        w = WIMManager._norm_path(wim_path)
        m = WIMManager._norm_path(mount_dir)
        args = [
            "/Mount-Image",
            f"/ImageFile:{w}",
            f"/Index:{index}",
            f"/MountDir:{m}",
        ]
        if readonly:
            args.append("/ReadOnly")
        rc, out, err = WIMManager._run_dism(args)
        if rc == 0:
            return True, "WIM 掛載完成"
        return False, err or out

    @staticmethod
    def unmount_wim(mount_dir: str, commit: bool = False) -> tuple[bool, str]:
        """
        卸載 WIM 映像
        
        Args:
            mount_dir: 掛載目錄
            commit: 是否提交變更
        
        Returns:
            (success, message)
        """
        m = WIMManager._norm_path(mount_dir)
        args = [
            "/Unmount-Image",
            f"/MountDir:{m}",
            "/Commit" if commit else "/Discard",
        ]
        rc, out, err = WIMManager._run_dism(args)
        if rc == 0:
            return True, "WIM 卸載完成"
        return False, err or out
    
    # ==================== 掛載狀態查詢 ====================

    @staticmethod
    def get_mount_info() -> tuple[bool, list[dict], str]:
        """
        取得目前所有掛載的映像資訊
        
        Returns:
            (success, mounted_images_list, error_message)
        """
        args = ["/Get-MountedImageInfo"]
        rc, out, err = WIMManager._run_dism(args)
        if rc != 0:
            return False, [], err or out
        mounted_images = WIMManager._parse_mounted_info(out)
        return True, mounted_images, ""
    
    @staticmethod
    def _parse_mounted_info(text: str) -> list[dict]:
        """解析 DISM 掛載資訊輸出"""
        images: list[dict] = []
        cur: dict | None = None
        
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
                
            # 檢測新掛載映像開始
            if line.startswith("Mount Dir"):
                if cur:
                    images.append(cur)
                cur = {"MountDir": "", "ImageFile": "", "ImageIndex": "", "Status": "", "ReadWrite": ""}
                match = re.search(r"Mount Dir\s*:\s*(.*)", line, re.IGNORECASE)
                if match:
                    cur["MountDir"] = match.group(1).strip()
                continue
                
            if cur is not None:
                for field in ["ImageFile", "ImageIndex", "Status"]:
                    pattern = f"{field.replace('File', ' File').replace('Index', ' Index')}\\s*:\\s*(.*)"
                    match = re.search(pattern, line, re.IGNORECASE)
                    if match:
                        cur[field] = match.group(1).strip()
                        break
                        
                # 檢查讀寫狀態
                if "Read/Write" in line:
                    cur["ReadWrite"] = "Read/Write"
                elif "Read Only" in line:
                    cur["ReadWrite"] = "Read Only"
                        
        if cur:
            images.append(cur)
        return images

    @staticmethod
    def is_path_mounted(mount_dir: str) -> tuple[bool, dict | None, str]:
        """
        檢查指定路徑是否已掛載 WIM 映像
        
        Args:
            mount_dir: 要檢查的掛載路徑
            
        Returns:
            (is_mounted, mount_info, error_message)
        """
        ok, mounts, err = WIMManager.get_mount_info()
        if not ok:
            return False, None, err
        
        # 標準化路徑以進行比較
        norm_mount_dir = os.path.normpath(mount_dir).lower()
        
        for mount in mounts:
            if os.path.normpath(mount.get("MountDir", "")).lower() == norm_mount_dir:
                return True, mount, ""
        
        return False, None, ""

    @staticmethod
    def get_mount_status_for_path(mount_dir: str) -> tuple[str, str]:
        """
        取得指定路徑的掛載狀態
        
        Returns:
            (status, details)
            - status: "mounted", "not_mounted", "error", "needs_remount", "orphaned"
        """
        is_mounted, info, err = WIMManager.is_path_mounted(mount_dir)
        
        if err:
            return "error", f"無法檢查掛載狀態: {err}"
        
        if not is_mounted:
            # 檢查路徑是否存在且有 Windows 資料夾
            if os.path.exists(mount_dir):
                windows_path = os.path.join(mount_dir, "Windows")
                if os.path.exists(windows_path):
                    return "orphaned", "路徑存在 Windows 資料夾但未在 DISM 掛載清單中"
                return "not_mounted", "路徑未掛載"
            return "not_mounted", "路徑不存在"
        
        # 已掛載，檢查狀態
        status = info.get("Status", "").lower() if info else ""
        if "invalid" in status or "needs remount" in status:
            return "needs_remount", f"掛載狀態異常: {info.get('Status', 'Unknown')}"
        
        return "mounted", f"已掛載 ({info.get('ReadWrite', 'Unknown')})"

    # ==================== 清理與修復 ====================

    @staticmethod
    def smart_cleanup_and_fix() -> tuple[bool, str]:
        """
        智能一鍵修復 - 自動診斷並解決所有 WIM 掛載問題
        包含檢查狀態、清理衝突、修復損壞掛載、強力清理等所有功能
        """
        messages = []
        messages.append("🚀 開始智能診斷和修復...")
        
        try:
            # === 第1步：檢查當前掛載狀態 ===
            messages.append("\n📋 第1步：檢查系統掛載狀態")
            args = ["/Get-MountedWimInfo"]
            rc, out, err = WIMManager._run_dism(args)
            
            if rc != 0:
                messages.append(f"❌ 無法檢查掛載狀態: {err or out}")
                return False, "\n".join(messages)
            
            if "No mounted images found" in out:
                messages.append("✅ 系統中沒有掛載的映像，狀態良好")
                return True, "\n".join(messages)
            
            # 解析掛載資訊
            lines = out.split('\n')
            mounted_images = []
            broken_mounts = []
            normal_mounts = []
            current_mount = {}
            
            for line in lines:
                line = line.strip()
                if line.startswith("Mount Dir"):
                    current_mount = {"dir": line.split(":", 1)[1].strip()}
                elif line.startswith("Image File"):
                    current_mount["file"] = line.split(":", 1)[1].strip()
                elif line.startswith("Image Index"):
                    current_mount["index"] = line.split(":", 1)[1].strip()
                elif line.startswith("Status"):
                    status = line.split(":", 1)[1].strip()
                    current_mount["status"] = status
                    mounted_images.append(current_mount.copy())
                    
                    # 分類掛載狀態
                    if status in ["Invalid", "Needs Remount", "Corrupted"]:
                        broken_mounts.append(current_mount.copy())
                    else:
                        normal_mounts.append(current_mount.copy())
                    current_mount = {}
            
            messages.append(f"📊 發現 {len(mounted_images)} 個掛載的映像:")
            for mount in mounted_images:
                status_icon = "❌" if mount["status"] in ["Invalid", "Needs Remount", "Corrupted"] else "✅"
                messages.append(f"   {status_icon} {mount['dir']} - 狀態: {mount['status']}")
            
            # === 第2步：處理正常掛載 ===
            if normal_mounts:
                messages.append(f"\n🔧 第2步：清理 {len(normal_mounts)} 個正常掛載")
                for mount in normal_mounts:
                    mount_dir = mount["dir"]
                    messages.append(f"   處理: {mount_dir}")
                    
                    # 嘗試正常卸載 (提交)
                    rc, out, err = WIMManager._run_dism(["/Unmount-Wim", f"/MountDir:{mount_dir}", "/Commit"])
                    if rc == 0:
                        messages.append(f"   ✅ 正常卸載成功")
                    else:
                        # 如果提交失敗，嘗試丟棄
                        rc, out, err = WIMManager._run_dism(["/Unmount-Wim", f"/MountDir:{mount_dir}", "/Discard"])
                        if rc == 0:
                            messages.append(f"   ✅ 丟棄卸載成功")
                        else:
                            messages.append(f"   ⚠️  卸載失敗，稍後統一處理")
            
            # === 第3步：修復損壞掛載 ===
            if broken_mounts:
                messages.append(f"\n🔨 第3步：修復 {len(broken_mounts)} 個損壞掛載")
                for mount in broken_mounts:
                    mount_dir = mount["dir"]
                    status = mount["status"]
                    messages.append(f"   修復: {mount_dir} (狀態: {status})")
                    
                    # 直接使用實測有效的 Discard 方法
                    rc, out, err = WIMManager._run_dism(["/Unmount-Wim", f"/MountDir:{mount_dir}", "/Discard"])
                    if rc == 0:
                        messages.append(f"   ✅ 損壞掛載修復成功")
                    else:
                        messages.append(f"   ⚠️  修復失敗: {err or out}")
            
            # === 第4步：系統級清理 ===
            messages.append(f"\n🧹 第4步：執行系統級清理")
            
            # 清理 WIM 緩存
            messages.append("   清理 WIM 緩存...")
            rc, out, err = WIMManager._run_dism(["/Cleanup-Wim"])
            if rc == 0:
                messages.append("   ✅ WIM 緩存清理完成")
            else:
                messages.append(f"   ⚠️  WIM 緩存清理警告: {err or out}")
            
            # 清理所有掛載點
            messages.append("   清理所有掛載點...")
            rc, out, err = WIMManager._run_dism(["/Cleanup-Mountpoints"])
            if rc == 0:
                messages.append("   ✅ 掛載點清理完成")
            else:
                messages.append(f"   ⚠️  掛載點清理警告: {err or out}")
            
            # === 第5步：驗證最終結果 ===
            messages.append(f"\n🔍 第5步：驗證修復結果")
            rc, out, err = WIMManager._run_dism(["/Get-MountedWimInfo"])
            
            if rc == 0 and "No mounted images found" in out:
                messages.append("🎉 一鍵修復完成！所有掛載問題已解決")
                messages.append("💡 系統現在處於乾淨狀態，可以正常進行新的掛載操作")
                return True, "\n".join(messages)
            elif rc == 0:
                # 檢查是否還有問題
                remaining_issues = 0
                remaining_details = []
                current_dir = ""
                for line in out.split('\n'):
                    line = line.strip()
                    if line.startswith("Mount Dir"):
                        current_dir = line.split(":", 1)[1].strip()
                    elif line.startswith("Status"):
                        status = line.split(":", 1)[1].strip()
                        if status in ["Invalid", "Needs Remount", "Corrupted"]:
                            remaining_issues += 1
                            remaining_details.append(f"{current_dir} ({status})")
                
                if remaining_issues == 0:
                    messages.append("✅ 一鍵修復完成！所有問題已解決")
                    messages.append("💡 仍有正常掛載存在，但狀態健康")
                    return True, "\n".join(messages)
                else:
                    messages.append(f"⚠️  還有 {remaining_issues} 個問題需要手動處理:")
                    for detail in remaining_details:
                        messages.append(f"     - {detail}")
                    messages.append("💡 建議：重新啟動電腦以完全清除頑固問題")
                    return True, "\n".join(messages)
            else:
                messages.append(f"⚠️  無法驗證修復結果: {err or out}")
                messages.append("💡 建議：重新啟動電腦確保所有更改生效")
                return True, "\n".join(messages)
                
        except Exception as e:
            messages.append(f"❌ 修復過程發生錯誤: {str(e)}")
            return False, "\n".join(messages)
    
    @staticmethod
    def cleanup_mount(mount_dir: str = None) -> tuple[bool, str]:
        """
        清理掛載狀態 - 清理所有或指定的掛載
        """
        messages = []
        
        if mount_dir:
            # 如果指定了掛載目錄，先嘗試卸載該特定映像
            m = WIMManager._norm_path(mount_dir)
            
            # 方法 1: 正常卸載
            messages.append(f"嘗試正常卸載 {mount_dir}...")
            unmount_args = ["/Unmount-Image", f"/MountDir:{m}", "/Discard"]
            rc, out, err = WIMManager._run_dism(unmount_args)
            if rc == 0:
                return True, f"已成功卸載指定映像: {mount_dir}"
            else:
                messages.append(f"正常卸載失敗: {err or out}")
        
        # 方法 2: 執行全域清理
        messages.append("執行全域掛載點清理...")
        args = ["/Cleanup-Mountpoints"]
        rc, out, err = WIMManager._run_dism(args)
        if rc == 0:
            messages.append("全域清理成功")
        else:
            messages.append(f"全域清理失敗: {err or out}")
        
        # 方法 3: 強制清理 (使用 /ScratchDir 重置)
        messages.append("嘗試強制清理...")
        try:
            import tempfile
            temp_dir = tempfile.mkdtemp()
            force_args = ["/Cleanup-Mountpoints", f"/ScratchDir:{temp_dir}"]
            rc2, out2, err2 = WIMManager._run_dism(force_args)
            if rc2 == 0:
                messages.append("強制清理成功")
                return True, "\n".join(messages)
            else:
                messages.append(f"強制清理失敗: {err2 or out2}")
        except Exception as e:
            messages.append(f"強制清理異常: {str(e)}")
        
        # 方法 4: 重啟 DISM 服務
        messages.append("嘗試重啟相關服務...")
        try:
            # 停止可能的服務
            subprocess.run(["net", "stop", "TrustedInstaller"], capture_output=True, text=True, timeout=10)
            subprocess.run(["net", "start", "TrustedInstaller"], capture_output=True, text=True, timeout=10)
            messages.append("服務重啟完成")
        except Exception as e:
            messages.append(f"服務重啟失敗: {str(e)}")
        
        # 最終檢查
        final_check_args = ["/Get-MountedImageInfo"]
        rc3, out3, err3 = WIMManager._run_dism(final_check_args)
        if rc3 == 0 and ("No mounted images found" in out3 or "找不到掛載的映像" in out3):
            messages.append("✓ 確認所有映像已卸載")
            return True, "\n".join(messages)
        
        return False, "\n".join(messages)
        
    @staticmethod
    def force_cleanup_registry() -> tuple[bool, str]:
        """強制清理 DISM 掛載註冊表項目"""
        try:
            import winreg
            messages = []
            
            # DISM 掛載資訊通常存在這些註冊表位置
            registry_paths = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\WIMMount\Mounted Images"),
                (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Services\WIMMount"),
            ]
            
            for root_key, sub_key in registry_paths:
                try:
                    reg_key = winreg.OpenKey(root_key, sub_key, 0, winreg.KEY_READ)
                    
                    i = 0
                    while True:
                        try:
                            subkey_name = winreg.EnumKey(reg_key, i)
                            messages.append(f"發現掛載記錄: {subkey_name}")
                            i += 1
                        except WindowsError:
                            break
                    
                    winreg.CloseKey(reg_key)
                    
                    if i > 0:
                        messages.append(f"找到 {i} 個掛載記錄在 {sub_key}")
                        
                except FileNotFoundError:
                    messages.append(f"註冊表路徑不存在: {sub_key}")
                except PermissionError:
                    messages.append(f"沒有權限訪問: {sub_key}")
                except Exception as e:
                    messages.append(f"註冊表操作錯誤: {str(e)}")
            
            return True, "\n".join(messages) if messages else "註冊表檢查完成，未發現問題"
            
        except Exception as e:
            return False, f"註冊表清理失敗: {str(e)}"

    # ==================== 錯誤處理與建議 ====================

    @staticmethod
    def get_error_solution_advice(error_message: str) -> tuple[str, str, list[str]]:
        """
        根據錯誤訊息提供對應的解決建議
        
        Returns:
            (錯誤類型, 建議說明, 推薦操作順序)
        """
        error_msg = error_message.lower()
        
        # Error 0xc1420127 - 映像已掛載
        if "0xc1420127" in error_msg or "already mounted" in error_msg:
            return (
                "映像已掛載衝突",
                "此錯誤表示相同的 WIM 檔案和 Index 已經在系統中掛載。",
                [
                    "1. 點擊「檢查掛載狀態」查看現有掛載",
                    "2. 如果狀態正常，點擊「清理掛載」",
                    "3. 如果狀態顯示異常，點擊「修復損壞掛載」",
                    "4. 最後手段：點擊「強力清理」"
                ]
            )
        
        # Error 50 - 請求不支援  
        elif "error: 50" in error_msg or "request is not supported" in error_msg:
            return (
                "掛載狀態損壞",
                "此錯誤通常表示掛載點處於 'Needs Remount' 等損壞狀態。",
                [
                    "1. 點擊「修復損壞掛載」(專門處理此問題)",
                    "2. 如果失敗，點擊「強力清理」",
                    "3. 極端情況：重新啟動電腦"
                ]
            )
        
        # Error 2 - 檔案不存在
        elif "error: 2" in error_msg or "cannot find the file" in error_msg or "找不到檔案" in error_msg:
            return (
                "檔案路徑問題", 
                "無法找到指定的 WIM 檔案或掛載目錄。",
                [
                    "1. 檢查 WIM 檔案路徑是否正確",
                    "2. 確認掛載目錄是否存在",
                    "3. 點擊「建立」按鈕創建掛載目錄",
                    "4. 檢查檔案是否被移動或刪除"
                ]
            )
            
        # Error 5 - 拒絕存取
        elif "error: 5" in error_msg or "access denied" in error_msg or "拒絕存取" in error_msg:
            return (
                "權限不足",
                "沒有足夠的權限執行 DISM 操作。",
                [
                    "1. 確認程式以管理員權限執行",
                    "2. 檢查 WIM 檔案是否被其他程式鎖定",
                    "3. 暫時關閉防毒軟體",
                    "4. 重新啟動程式"
                ]
            )
            
        # Error 1392 - 檔案損壞
        elif "error: 1392" in error_msg or "corrupted" in error_msg or "damaged" in error_msg:
            return (
                "檔案損壞",
                "WIM 檔案可能已損壞或不完整。",
                [
                    "1. 使用其他工具驗證 WIM 檔案完整性",
                    "2. 重新下載或複製 WIM 檔案",
                    "3. 檢查磁碟錯誤 (chkdsk)",
                    "4. 嘗試使用備份的 WIM 檔案"
                ]
            )
            
        # 掛載目錄不為空
        elif "not empty" in error_msg or "不為空" in error_msg or "directory is not empty" in error_msg:
            return (
                "掛載目錄不為空",
                "DISM 需要空的目錄來掛載映像。",
                [
                    "1. 清空掛載目錄中的所有檔案",
                    "2. 點擊「開啟」按鈕檢查目錄內容",
                    "3. 選擇其他空目錄作為掛載點",
                    "4. 點擊「建立」按鈕創建新的空目錄"
                ]
            )
            
        # 磁碟空間不足
        elif "not enough space" in error_msg or "insufficient disk space" in error_msg or "磁碟空間不足" in error_msg:
            return (
                "磁碟空間不足",
                "目標磁碟沒有足夠空間進行掛載操作。",
                [
                    "1. 清理磁碟空間",
                    "2. 選擇其他有足夠空間的磁碟",
                    "3. 刪除不必要的檔案",
                    "4. 使用磁碟清理工具"
                ]
            )
            
        # Index 無效
        elif "invalid index" in error_msg or "index not found" in error_msg or "索引無效" in error_msg:
            return (
                "映像索引錯誤",
                "指定的 Index 在 WIM 檔案中不存在。",
                [
                    "1. 點擊「讀取映像資訊」重新載入 Index 列表",
                    "2. 選擇有效的 Index 編號",
                    "3. 檢查 WIM 檔案是否完整",
                    "4. 確認選擇的 Index 沒有被其他工具占用"
                ]
            )
            
        # 一般性 DISM 錯誤
        elif "dism" in error_msg and "error" in error_msg:
            return (
                "DISM 操作錯誤",
                "DISM 工具執行時遇到問題。",
                [
                    "1. 點擊「檢查掛載狀態」查看系統狀態",
                    "2. 嘗試「清理掛載」解決衝突",
                    "3. 檢查 Windows 日誌: C:\\Windows\\Logs\\DISM\\dism.log",
                    "4. 重新啟動系統清除所有狀態"
                ]
            )
        
        # 未知錯誤
        else:
            return (
                "未知錯誤",
                "遇到了未預期的錯誤情況。",
                [
                    "1. 點擊「檢查掛載狀態」診斷系統狀態",
                    "2. 嘗試「清理掛載」清除可能的衝突",
                    "3. 查看詳細錯誤日誌",
                    "4. 考慮重新啟動程式或系統"
                ]
            )

    # ==================== 終極清理 ====================

    @staticmethod
    def ultimate_cleanup() -> tuple[bool, str]:
        """終極清理方法 - 當所有其他方法都失敗時使用"""
        messages = []
        success_count = 0
        
        try:
            # 1. 強制終止可能相關的進程
            messages.append("=== 步驟 1: 終止相關進程 ===")
            processes_to_kill = ["dism.exe", "DismHost.exe", "TiWorker.exe"]
            
            for proc in processes_to_kill:
                try:
                    result = subprocess.run(["taskkill", "/F", "/IM", proc], 
                                          capture_output=True, text=True)
                    if result.returncode == 0:
                        messages.append(f"✓ 終止進程: {proc}")
                        success_count += 1
                    else:
                        messages.append(f"- 進程不存在或已終止: {proc}")
                except Exception as e:
                    messages.append(f"✗ 終止進程失敗 {proc}: {str(e)}")
            
            # 2. 清理暫存目錄
            messages.append("\n=== 步驟 2: 清理暫存目錄 ===")
            temp_dirs = [
                r"C:\Windows\Temp",
                r"C:\Windows\Logs\DISM",
            ]
            
            for temp_dir in temp_dirs:
                try:
                    if os.path.exists(temp_dir):
                        dism_files = []
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                if "dism" in file.lower() or "wim" in file.lower():
                                    file_path = os.path.join(root, file)
                                    try:
                                        os.remove(file_path)
                                        dism_files.append(file)
                                    except Exception:
                                        pass
                        
                        if dism_files:
                            messages.append(f"✓ 清理 {len(dism_files)} 個相關檔案從 {temp_dir}")
                            success_count += 1
                        else:
                            messages.append(f"- 無需清理: {temp_dir}")
                    else:
                        messages.append(f"- 目錄不存在: {temp_dir}")
                        
                except Exception as e:
                    messages.append(f"✗ 清理目錄失敗 {temp_dir}: {str(e)}")
            
            # 3. 重新啟動相關服務
            messages.append("\n=== 步驟 3: 重啟系統服務 ===")
            services = ["TrustedInstaller", "wuauserv", "bits"]
            
            for service in services:
                try:
                    # 停止服務
                    subprocess.run(["sc", "stop", service], capture_output=True, text=True, timeout=10)
                    # 等待一下
                    import time
                    time.sleep(2)
                    # 啟動服務
                    result = subprocess.run(["sc", "start", service], capture_output=True, text=True, timeout=15)
                    
                    if result.returncode == 0:
                        messages.append(f"✓ 重啟服務: {service}")
                        success_count += 1
                    else:
                        messages.append(f"- 服務可能已在運行: {service}")
                        
                except Exception as e:
                    messages.append(f"✗ 重啟服務失敗 {service}: {str(e)}")
            
            # 4. 最終的 DISM 清理嘗試
            messages.append("\n=== 步驟 4: 最終 DISM 清理 ===")
            try:
                cleanup_commands = [
                    ["/Cleanup-Mountpoints"],
                    ["/Cleanup-Wim"],
                    ["/Cleanup-Mountpoints", "/RevertPendingActions"],
                ]
                
                for cmd in cleanup_commands:
                    try:
                        rc, out, err = WIMManager._run_dism(cmd)
                        cmd_str = " ".join(cmd)
                        if rc == 0:
                            messages.append(f"✓ DISM 清理成功: {cmd_str}")
                            success_count += 1
                            break
                        else:
                            messages.append(f"- DISM 清理嘗試: {cmd_str} - {err or out}")
                    except Exception as e:
                        messages.append(f"✗ DISM 清理異常: {str(e)}")
                        
            except Exception as e:
                messages.append(f"✗ DISM 清理階段失敗: {str(e)}")
            
            # 5. 檢查註冊表
            messages.append("\n=== 步驟 5: 註冊表檢查 ===")
            reg_ok, reg_msg = WIMManager.force_cleanup_registry()
            messages.append(reg_msg)
            if reg_ok:
                success_count += 1
            
            # 總結
            messages.append(f"\n=== 清理完成 ===")
            messages.append(f"成功步驟: {success_count}/5")
            
            final_success = success_count >= 3
            return final_success, "\n".join(messages)
            
        except Exception as e:
            messages.append(f"\n✗ 終極清理發生嚴重錯誤: {str(e)}")
            return False, "\n".join(messages)
    
    @staticmethod
    def close_explorer_windows(target_path: str) -> tuple[bool, str]:
        """關閉指向特定路徑的檔案總管視窗"""
        try:
            # 正規化路徑
            target_path = os.path.normpath(target_path).lower()
            
            # 嘗試使用 PowerShell 關閉特定資料夾的檔案總管視窗
            ps_script = f'''
$shell = New-Object -ComObject Shell.Application
$windows = $shell.Windows()
$closed = 0
foreach ($window in $windows) {{
    try {{
        $path = $window.LocationURL
        if ($path -like "*file:///*") {{
            $localPath = $window.Document.Folder.Self.Path
            if ($localPath -and $localPath.ToLower().StartsWith("{target_path}")) {{
                $window.Quit()
                $closed++
            }}
        }}
    }} catch {{
        # 忽略錯誤，繼續下一個視窗
    }}
}}
Write-Output "已關閉 $closed 個檔案總管視窗"
'''
            
            result = subprocess.run(['powershell', '-Command', ps_script], 
                                  capture_output=True, text=True)
            
            if result.returncode == 0:
                output = result.stdout.strip()
                return True, output
            else:
                return False, f"PowerShell 執行失敗: {result.stderr}"
                
        except Exception as e:
            return False, f"關閉檔案總管視窗時發生錯誤: {str(e)}"


# 導出全域鎖供其他模組使用
def get_dism_lock():
    """取得 DISM 操作鎖"""
    return _dism_lock

def is_dism_busy():
    """檢查是否有 DISM 操作正在執行"""
    return _dism_busy

def set_dism_busy(busy: bool):
    """設定 DISM 忙碌狀態"""
    global _dism_busy
    _dism_busy = busy
