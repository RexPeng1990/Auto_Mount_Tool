"""
DISM 命令執行器模組

負責底層 DISM 命令執行、進度回調、線程安全鎖管理
"""

import os
import re
import subprocess
import threading
from typing import Callable

# Type alias for progress callback
ProgressCallback = Callable[[str], None] | None

# 全域 DISM 操作鎖 - 確保同時只有一個 DISM 操作執行
_dism_lock = threading.Lock()
_dism_busy = False


def get_dism_lock() -> threading.Lock:
    """取得 DISM 操作鎖"""
    return _dism_lock


def is_dism_busy() -> bool:
    """檢查是否有 DISM 操作正在執行"""
    return _dism_busy


def set_dism_busy(busy: bool) -> None:
    """設定 DISM 忙碌狀態"""
    global _dism_busy
    _dism_busy = busy


def norm_path(path: str) -> str:
    """標準化路徑 - 移除尾部反斜線 (DISM 需要)"""
    return os.path.normpath(path).rstrip("\\")


def run_dism(args: list[str], timeout: int = 600) -> tuple[int, str, str]:
    """
    執行 DISM 命令並回傳結果
    
    Args:
        args: DISM 參數列表
        timeout: 超時秒數，預設 600 秒
        
    Returns:
        (return_code, stdout, stderr)
    """
    with _dism_lock:
        global _dism_busy
        _dism_busy = True
        try:
            cmd = ["dism.exe", "/English"] + args
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            return result.returncode, result.stdout, result.stderr
        except subprocess.TimeoutExpired:
            return -1, "", "DISM 操作超時"
        except Exception as e:
            return -1, "", str(e)
        finally:
            _dism_busy = False


def run_dism_with_progress(
    args: list[str],
    progress_callback: ProgressCallback = None,
    timeout: int = 1800
) -> tuple[int, str, str]:
    """
    執行 DISM 命令並支援即時進度回調
    
    Args:
        args: DISM 參數列表
        progress_callback: 進度回調函數，接收進度字串
        timeout: 超時秒數，預設 1800 秒
        
    Returns:
        (return_code, stdout, stderr)
    """
    with _dism_lock:
        global _dism_busy
        _dism_busy = True
        try:
            cmd = ["dism.exe", "/English"] + args
            
            # 使用 Popen 以便即時讀取輸出
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            
            stdout_lines = []
            stderr_lines = []
            
            # 讀取 stdout 並解析進度
            if process.stdout:
                for line in iter(process.stdout.readline, ''):
                    if not line:
                        break
                    stdout_lines.append(line)
                    
                    # 解析進度百分比
                    if progress_callback:
                        # 匹配進度格式如 "[ 45.0% ]" 或 "45.0 %"
                        progress_match = re.search(r'\[\s*(\d+(?:\.\d+)?)\s*%\s*\]|(\d+(?:\.\d+)?)\s*%', line)
                        if progress_match:
                            percent = progress_match.group(1) or progress_match.group(2)
                            progress_callback(f"{percent}%")
                        # 也處理狀態訊息
                        elif line.strip() and not line.startswith('==='):
                            progress_callback(line.strip()[:50])
            
            # 等待完成
            try:
                _, stderr = process.communicate(timeout=timeout)
                if stderr:
                    stderr_lines.append(stderr)
            except subprocess.TimeoutExpired:
                process.kill()
                return -1, "", "DISM 操作超時"
            
            return process.returncode, "".join(stdout_lines), "".join(stderr_lines)
            
        except Exception as e:
            return -1, "", str(e)
        finally:
            _dism_busy = False


def parse_dism_progress(line: str) -> str | None:
    """
    從 DISM 輸出行解析進度百分比
    
    Returns:
        進度字串如 "45.0%"，或 None 如果無法解析
    """
    progress_match = re.search(r'\[\s*(\d+(?:\.\d+)?)\s*%\s*\]|(\d+(?:\.\d+)?)\s*%', line)
    if progress_match:
        percent = progress_match.group(1) or progress_match.group(2)
        return f"{percent}%"
    return None
