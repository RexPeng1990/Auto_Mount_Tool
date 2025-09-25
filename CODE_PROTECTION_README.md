# WIM Driver Manager - 代碼保護說明

## 概述
本工具提供了多層次的代碼保護，防止基本的逆向工程和惡意分析。**現已改用 onedir 模式打包，確保更好的相容性和穩定性。**

## 保護版本

### 1. 直接打包版本 (最高穩定性)
- **目錄**: `direct_release/WIM_Driver_Manager_Direct/`
- **執行文件**: `WIM_Driver_Manager_Direct.exe`
- **大小**: ~2 MB (主程序) + 依賴庫
- **保護級別**: 無 (原始代碼)
- **特性**:
  - ✅ 原始代碼直接打包
  - ✅ 最高穩定性和相容性
  - ✅ onedir 目錄式打包
  - ✅ 所有標準庫完整包含

### 2. 簡化保護版本 (推薦)
- **目錄**: `simple_release/WIM_Driver_Manager_Simple/`
- **執行文件**: `WIM_Driver_Manager_Simple.exe`
- **大小**: ~2 MB (主程序) + 依賴庫
- **保護級別**: 基本
- **特性**:
  - ✅ 字符串編碼保護
  - ✅ 基本完整性檢查
  - ✅ onedir 目錄式打包
  - ✅ 高相容性

### 3. 進階保護版本
- **目錄**: `release/WIM_Driver_Manager_Protected/`
- **執行文件**: `WIM_Driver_Manager_Protected.exe`
- **大小**: ~2 MB (主程序) + 依賴庫
- **保護級別**: 中級
- **特性**:
  - ✅ 多層代碼加密
  - ✅ Base64 + zlib + Marshal 保護
  - ✅ 動態解密載入
  - ✅ 修復 configparser 模組問題

## 安全機制詳細說明

### 反調試保護
- **IsDebuggerPresent API**: 檢測是否有調試器附加
- **進程掃描**: 定期掃描系統中的可疑進程
- **時間檢查**: 監控代碼執行時間異常
- **自動退出**: 檢測到威脅時立即終止程序

### 代碼加密
1. **第一層**: XOR 加密使用64位隨機密鑰
2. **第二層**: Base64 編碼隱藏二進制特徵
3. **第三層**: zlib 壓縮減少體積和混淆
4. **第四層**: 最終 Base64 編碼存儲

### 運行時保護
- **動態解密**: 代碼在運行時才解密執行
- **內存保護**: 敏感數據不持久存儲
- **定時檢查**: 每30-180秒隨機檢查威脅
- **多線程監控**: 後台持續監控系統狀態

## 使用方法

### 執行直接打包版本 (最穩定)
```bash
# 管理員權限執行直接版本
powershell -Command "Start-Process '.\\direct_release\\WIM_Driver_Manager_Direct\\WIM_Driver_Manager_Direct.exe' -Verb RunAs"
```

### 執行簡化保護版本
```bash
# 管理員權限執行簡化版本
powershell -Command "Start-Process '.\\simple_release\\WIM_Driver_Manager_Simple\\WIM_Driver_Manager_Simple.exe' -Verb RunAs"
```

### 執行進階保護版本
```bash
# 管理員權限執行進階版本  
powershell -Command "Start-Process '.\\release\\WIM_Driver_Manager_Protected\\WIM_Driver_Manager_Protected.exe' -Verb RunAs"
```

### 重新構建版本
```bash
# 直接打包版本 (最穩定)
python direct_build.py

# 簡化保護版本
python simple_build.py

# 進階保護版本
python advanced_build.py
```

## 目錄式打包的優勢

### 相較於單文件模式
1. **啟動速度更快**: 無需解壓縮臨時文件
2. **相容性更好**: 避免防毒軟件誤報
3. **穩定性更高**: 減少執行時錯誤
4. **維護性更佳**: 可以單獨替換依賴庫

### 文件結構
```
WIM_Driver_Manager_Simple/
├── WIM_Driver_Manager_Simple.exe  (主執行文件)
└── _internal/                     (依賴庫目錄)
    ├── Python DLL files
    ├── tkinter 相關文件
    └── 其他依賴庫
```

## 注意事項

### 安全提醒
1. **保管好源代碼**: 加密密鑰嵌入在保護工具中，源代碼泄露將導致保護失效
2. **定期更新**: 逆向技術不斷發展，建議定期更新保護機制
3. **多重分發**: 可以生成不同密鑰的版本，避免批量破解

### 性能影響
- **啟動時間**: 增加1-3秒解密時間
- **內存使用**: 增加約5-10MB內存開銷
- **CPU使用**: 後台監控增加輕微CPU負載

### 兼容性
- **操作系統**: Windows 7/8/10/11 (64位)
- **權限要求**: 需要管理員權限（DISM操作必需）
- **依賴項**: 所有依賴已打包，無需額外安裝

## 防護等級評估

| 攻擊類型 | 基本版本 | 終極版本 | 說明 |
|---------|---------|----------|------|
| 靜態分析 | 🟡 中等 | 🟢 良好 | 代碼完全加密混淆 |
| 動態調試 | 🔴 弱 | 🟢 良好 | 主動反調試檢測 |
| 內存轉儲 | 🔴 弱 | 🟡 中等 | 運行時解密保護 |
| 網絡分析 | 🟢 良好 | 🟢 良好 | 無網絡通信 |
| 虛擬機分析 | 🔴 無 | 🟡 中等 | 基本虛擬機檢測 |

## 工具文件說明

- `main.py` - 原始源代碼
- `simple_build.py` - 簡化保護構建器 **(推薦使用)**
- `advanced_build.py` - 進階保護構建器
- `ultimate_build.py` - 終極保護構建器 (已停用)
- `security_protection.py` - 安全保護模組
- `build_config.spec` - PyInstaller 配置文件

## 版本比較

| 特性 | 簡化版本 | 進階版本 | 說明 |
|------|---------|----------|------|
| 啟動速度 | 🟢 快速 | 🟡 中等 | 簡化版無複雜解密 |
| 相容性 | 🟢 優秀 | 🟡 良好 | 目錄模式相容性更佳 |
| 保護強度 | 🟡 基本 | 🟢 中級 | 進階版多層加密 |
| 穩定性 | 🟢 優秀 | 🟢 良好 | 兩版本都已測試 |
| 文件大小 | 🟢 較小 | 🟡 中等 | 相差不大 |
| 維護難度 | 🟢 簡單 | 🟡 中等 | 簡化版更易維護 |

**建議**: 日常使用選擇簡化版本，如需更高保護等級選擇進階版本。

## 技術支援

如需更高級的保護或遇到問題，請：
1. 檢查Windows Defender等防毒軟件是否誤報
2. 確認以管理員權限運行
3. 查看事件日誌尋找錯誤信息

---
**免責聲明**: 本保護系統旨在防止一般性的逆向工程，對於專業的安全研究可能無法提供完全保護。請根據實際安全需求選擇合適的保護級別。
