# consolidate_specs — 多機種規格彙整工具

> 將多個機種規格 Excel 檔案（`.xlsx` / `.xls`）彙整為一份 Excel，方便比對與管理。

跟 Claude Code 的三段對話

```palintext 
Excel 檔案描述:
只有一個 Worksheet tab 叫 Specifications
A 欄位是空白
B1 是機種名稱 (model name)
B2 往下 B3,B4...會是規格名稱 (Example: CPU, Chipset...)
C2 往下 C3,C4...是對應規格的說明
但並非每個 B 欄位都有資訊, 舉例：
B21 的內容是 Memory
C21 = Dual Channel DDR5 Memory Technology
C22 = 4 x DDR5 DIMM Slots
有關 Memory 規格的描述一直到 C26 結束
B22 到 B26 都是空白
B27 就是下一個規格名稱 Expansion Slot
以上描述你是否了解 ？
 
有一堆這樣的 Excel 檔案，一個檔案就是一個機種.
開一個新的 Excel 檔案彙整所有機種:
新的 Excel 檔案 B1 = "Model", C1 = 該機種名稱,
將所有舊檔案依序處理到 C,D,E, ... 欄位
這個用 Python 可以做嗎 ? 做好我要交給 PM 使用.
困難在於每個舊檔案中相同的規格描述所佔的列數可能不一樣.
以上描述你是否了解 ？
 
請告訴我該把測試用的 Excel 舊檔案複製到哪個路徑.
請你檢視一下後再整理一下你的 actions.
```

## 功能說明

- **增量匯入**：偵測已存在的彙整檔，只加入新機種，不會重複匯入
- **自動跳過**：已匯入的機種會自動略過，並顯示提示
- **`--force` 重建**：加上 `--force` 參數可忽略既有檔案，從頭全部重建
- **格式化輸出**：自動套用粗體標題、文字自動換行、凍結窗格等格式

---

## 給開發者：從原始碼執行

### 環境需求

- Python 3.8 以上

### 安裝相依套件

```bash
pip install -r requirements.txt
```

### CLI 用法

```bash
python consolidate_specs.py <folder> [-o OUTPUT] [-f]
```

| 參數 | 說明 |
|------|------|
| `folder` | 包含規格檔案的資料夾路徑（必填） |
| `-o`, `--output` | 輸出檔案路徑（預設：`<folder>/consolidated_specs.xlsx`） |
| `-f`, `--force` | 忽略既有檔案，從頭全部重建 |

### 使用範例

```bash
# 基本用法：彙整 spec_files 資料夾
python consolidate_specs.py spec_files

# 指定輸出路徑
python consolidate_specs.py spec_files -o output/all_specs.xlsx

# 強制重建（忽略既有檔案）
python consolidate_specs.py spec_files --force
```

---

## 給開發者：打包成 Windows 執行檔

由於 PyInstaller 不支援跨平台編譯，需在 **Windows 機器**上執行打包。

### 使用 build.bat（推薦）

1. 確認 Windows 機器已安裝 **Python 3.8+**
2. 將整個專案資料夾複製到 Windows 機器
3. 雙擊執行 `build.bat`
4. 打包完成後，`consolidate_specs.exe` 會出現在專案根目錄

`build.bat` 會自動完成：建立虛擬環境 → 安裝套件 → PyInstaller 打包 → 複製 `.exe` 到根目錄。

### 手動打包（備用）

```cmd
python -m venv venv
call venv\Scripts\activate.bat
pip install -r requirements.txt
pyinstaller --onefile --console --name consolidate_specs ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    consolidate_specs.py
copy dist\consolidate_specs.exe .
```

---

## 給 PM：安裝 VSCode + Python 直接執行

如果你願意在 Windows 上安裝 VSCode 和 Python，就能直接執行 `.py` 原始碼，不需要打包 `.exe`，也更方便日後更新腳本。

### Step 1：安裝 Python

1. 前往 [python.org](https://www.python.org/downloads/) 下載 **Python 3.8 以上**的 Windows installer
2. 執行安裝程式時，**務必勾選「Add Python to PATH」**（這是最關鍵的一步！）
3. 安裝完成後，開啟 CMD（命令提示字元），輸入以下指令確認安裝成功：

```cmd
python --version
```

看到類似 `Python 3.12.x` 的輸出就代表安裝成功。

### Step 2：安裝 VSCode

1. 前往 [code.visualstudio.com](https://code.visualstudio.com/) 下載並安裝 VSCode
2. 開啟 VSCode，點選左側「Extensions」圖示（四個方塊的圖示）
3. 搜尋 **Python**，安裝由 **Microsoft** 發行的官方 Python 擴充套件

### Step 3：安裝相依套件

1. 在 VSCode 中開啟專案資料夾（File → Open Folder）
2. 開啟終端機（Terminal → New Terminal）
3. 在終端機中輸入：

```cmd
pip install -r requirements.txt
```

### Step 4：執行程式

**方式一：在終端機輸入指令**

在 VSCode 的終端機中輸入：

```cmd
python consolidate_specs.py spec_files
```

**方式二：右鍵執行**

在 VSCode 中開啟 `consolidate_specs.py`，按右鍵 → **Run Python File**。

> 注意：右鍵執行不會帶入參數，適合用於測試。正式執行建議使用終端機指令。

### 這個方式的好處

- **不需打包 `.exe`**：腳本更新後直接可用，不需要重新打包
- **VSCode 整合**：提供語法高亮、終端機整合，操作更方便
- **透明度高**：可直接看到程式碼與輸出，方便排查問題

---

## 給 PM：使用 .exe 執行

不需要安裝 Python，只需要 `consolidate_specs.exe` 即可使用。

### 方法一：拖放執行（推薦）

1. 將 `run.bat` 和 `consolidate_specs.exe` 放在同一個資料夾
2. 將 `spec_files` 資料夾**拖放到 `run.bat` 上**
3. 程式會自動執行，完成後按任意鍵關閉視窗

### 方法二：開啟 CMD 手動輸入

```cmd
consolidate_specs.exe spec_files
```

#### 參數說明

| 參數 | 說明 |
|------|------|
| `--force` | 全部重建（忽略已匯入的機種） |
| `-o 路徑` | 指定輸出檔案路徑 |

```cmd
REM 強制重建
consolidate_specs.exe spec_files --force

REM 指定輸出位置
consolidate_specs.exe spec_files -o D:\output\all_specs.xlsx
```

---

## 檔案結構說明

```
專案目錄/
├── consolidate_specs.py   # Python 原始碼
├── requirements.txt       # Python 相依套件
├── build.bat              # Windows 打包腳本（開發者用）
├── run.bat                # 拖放啟動腳本（PM 用）
├── consolidate_specs.exe  # 打包後的執行檔（打包後產生）
└── README.md              # 本文件
```

---

## 常見問題 FAQ

### Windows Defender 跳出警告怎麼辦？

PyInstaller 打包的 `.exe` 可能被 Windows Defender 誤判為不明程式。

1. 點選「**更多資訊**」
2. 點選「**仍要執行**」

這是因為程式沒有數位簽章，屬於正常現象。

### 新增機種後怎麼用？

將新的規格檔案放入 `spec_files` 資料夾，重新執行即可。程式會自動偵測並只匯入新增的機種。

### 想全部重來怎麼辦？

加上 `--force` 參數即可從頭重建：

- **拖放方式**：開啟 CMD，輸入 `consolidate_specs.exe spec_files --force`
- **原始碼方式**：`python consolidate_specs.py spec_files --force`
