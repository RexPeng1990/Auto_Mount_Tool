# WIM Driver Manager - 打包保護系統

## 概述

本工具提供了一套完整的代碼打包保護系統，支援多種保護級別，從無保護的直接打包到帶有反調試功能的終極保護。

## 快速開始

### 使用批處理腳本 (推薦)

```batch
# 運行交互式打包菜單
build.bat
```

### 使用 Python 腳本

```bash
# 構建簡單保護版本 (預設)
python build_all.py

# 構建特定保護級別
python build_all.py -l direct    # 直接打包
python build_all.py -l simple    # 簡單保護
python build_all.py -l advanced  # 進階保護
python build_all.py -l ultimate  # 終極保護

# 構建所有版本
python build_all.py -l all

# 清理並重新構建
python build_all.py -l all -c

# 構建並更新版本號
python build_all.py -l simple -b patch  # 增加修訂版本號
python build_all.py -l simple -b minor  # 增加次版本號
python build_all.py -l simple -b major  # 增加主版本號
```

## 保護級別

### 1. Direct (直接打包)
- **保護級別**: 無
- **輸出目錄**: `direct_release/WIM_Driver_Manager_Direct/`
- **特點**:
  - ✅ 原始代碼直接打包
  - ✅ 最高穩定性和相容性
  - ✅ 適合開發測試
  - ❌ 無代碼保護

### 2. Simple (簡單保護) - 推薦
- **保護級別**: 基本
- **輸出目錄**: `simple_release/WIM_Driver_Manager_Simple/`
- **特點**:
  - ✅ 字符串混淆
  - ✅ 環境檢查
  - ✅ 高相容性
  - ✅ 適合日常使用

### 3. Advanced (進階保護)
- **保護級別**: 中級
- **輸出目錄**: `release/WIM_Driver_Manager_Protected/`
- **特點**:
  - ✅ 多層加密 (XOR + zlib)
  - ✅ 動態解密載入
  - ✅ 適合正式發布

### 4. Ultimate (終極保護)
- **保護級別**: 最高
- **輸出目錄**: `ultimate_release/WIM_Driver_Manager_Ultimate/`
- **特點**:
  - ✅ 多層加密 (RC4 + zlib + 打散)
  - ✅ 反調試保護
  - ✅ 進程監控
  - ✅ 時間檢測
  - ✅ 最高安全性

## 保護機制詳解

### 加密層次
```
原始代碼
    ↓
[Level 1] 字符串混淆 - Base64 編碼敏感字符串
    ↓
[Level 2] XOR 加密 - 使用隨機密鑰加密
    ↓
[Level 3] RC4 加密 - 更強的流加密算法
    ↓
[Level 3] 位元組打散 - 打亂數據順序
    ↓
zlib 壓縮 - 減小體積並混淆
    ↓
Base64 編碼 - 最終輸出
```

### 反調試機制

1. **API 檢測**
   - `IsDebuggerPresent`: 檢測本地調試器
   - `CheckRemoteDebuggerPresent`: 檢測遠程調試器

2. **進程掃描**
   - 監控系統中的可疑進程
   - 檢測 IDA、OllyDbg、x64dbg 等

3. **時間檢測**
   - 監控代碼執行時間
   - 調試會導致執行變慢

4. **隨機檢查**
   - 檢查間隔隨機化
   - 增加繞過難度

## 文件結構

```
Auto_Mount_Tool/
├── build_all.py           # 統一打包腳本
├── build.bat              # 交互式打包批處理
├── protection_core.py     # 保護核心模組
├── version.txt            # 版本號文件
├── main.py                # 主程序源代碼
├── app/                   # 應用模組
│   ├── config.py
│   ├── driver_manager.py
│   ├── utils.py
│   └── wim_manager.py
│
├── direct_release/        # 直接打包輸出
├── simple_release/        # 簡單保護輸出
├── release/               # 進階保護輸出
└── ultimate_release/      # 終極保護輸出
```

## 依賴項

- Python 3.9+
- PyInstaller (`pip install pyinstaller`)

## 運行打包後的程序

所有打包版本都需要**管理員權限**運行：

```batch
# 方法 1: 右鍵選擇「以系統管理員身分執行」

# 方法 2: 使用 PowerShell
powershell -Command "Start-Process '.\simple_release\WIM_Driver_Manager_Simple\WIM_Driver_Manager_Simple.exe' -Verb RunAs"
```

## 自定義保護

使用 `protection_core.py` 可以自定義保護：

```python
from protection_core import create_protected_module

# 讀取源代碼
with open('main.py', 'r', encoding='utf-8') as f:
    source = f.read()

# 創建保護版本
protected = create_protected_module(
    source,
    protection_level=3,      # 1-3
    add_anti_debug=True,     # 添加反調試
)

# 保存
with open('protected_main.py', 'w', encoding='utf-8') as f:
    f.write(protected)
```

## 注意事項

1. **防毒軟件**: 加密和反調試代碼可能被防毒軟件誤報，建議添加白名單

2. **性能影響**: 更高的保護級別會略微增加啟動時間

3. **調試限制**: Ultimate 級別會主動檢測調試器，不適合開發環境

4. **版本管理**: 建議使用 `-b` 參數在發布時更新版本號

## 故障排除

### 打包失敗
- 確認已安裝 PyInstaller: `pip install pyinstaller`
- 檢查 Python 環境是否正確

### 程序無法啟動
- 嘗試使用 Direct 版本排除保護相關問題
- 檢查是否有缺少的依賴庫

### 被防毒軟件攔截
- 將輸出目錄添加到防毒軟件白名單
- 使用 Simple 級別替代 Ultimate

## 更新日誌

### v1.0.0
- 初始版本
- 支援 4 種保護級別
- 統一打包腳本

---

*本打包系統僅供代碼保護使用，請勿用於惡意目的。*
