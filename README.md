# 目的
將散落在各個地方的文件，通過拖拽或者SendTo，傳送到預設類別和命名規則的置頂文件夾。
<br>有文件清單檢查文件是否齊全、文件數量。
<br>可撤回/重做，關閉窗口後記憶就無了。

# 功能
**衹支援windows系統**
<br>**Python**
<br>**文件分類：** 支援透過按鈕或拖曳，將文件分配到預設類別資料夾。符號【】來框主分類名。歸檔邏輯是按照年月→文件列別。
<br>**清單視圖：** 右側樹狀結構顯示文件類別，並統計每類文件數量。
<br>**撤回/重做：** 支援逐步撤回與重做操作，保證操作可逆。
<br>**快捷键支援：**
<br>-`F5`刷新清單（PS.一般個清單自己會刷新）
<br>-`Ctrl+Z`撤回一步
<br>-`Ctrl+Y`重做一步
<br>**只開一個窗口：** `kinder_classify(主程序)`和`checklist_viewer(文件清單檢查)`都通過本地TCP和windows互斥量。`SendTo`時，文件列表只追加到現有窗口，不打開新窗口。
<br>**底部狀態欄提醒：** 使用和測試時候覺得彈窗好煩人，改為狀態欄通知。
<br>**Log日誌** 記錄動作與異常，可追蹤。

# 不足
因為懶，所以可能不利於別人二次開發和修改。

# 操作
## 1、下載好文件
<br>三個files
| 檔名                     | 說明              |
| ---------------------- | --------------- |
| `kinder_classify.py`   | 主程式             |
| `checklist_viewer.pyw` | 清單獨立視窗（可選）      |
| `config.json`          | 設定文件（類別、路徑等）    |
| `kinder_classify.log`  | 程式自動生成的日誌（可以忽略） |

## 2、確定好要放的位置，唔可以刪~
## 3、安裝python
①官網下載`python.org`
<br>②檢查下咗未：WIN+R→`cmd`→Enter
```python
python -V
```
③查Python位置
```python
where python
```
## 4、安裝Tkinter拖拽
```
pip install tkinterdnd2
```
測試Tkinter和Tkinterdnd2可不可用:WIN+R→cmd→Enter→輸入：
```
python -c "import tkinter; print('tk ok'); import tkinterdnd2; print('tkinterdnd2 ok')"
```
顯示如下，就可用：
```tk ok
tkinterdnd2 ok
```
## 5、測試Kinder_Classify和Checklist_viewer能不能用
雙擊打開，測試功能。

## 6、更改目標文件和分類名字
①目標路徑修改：
| 代碼                                   | 解釋                                    |
| ------------------------------------ | --------------------------------------- |
| `"out_root":`        | 根目錄。所有分類的目標資料夾都會以這個路徑為起點建立。                         |
| `"default_path_template":`        | 全局路徑                           |
|----|下面解釋他們關係|
| 有 `path_template` （在分類底下）| 使用此分類自己的目錄模板。                           |
| 沒有 `path_template` 但有 `dest_subdir`  | 使用全域 `default_path_template`，並在最後拼上子目錄。 |
| 沒有 `path_template` 也沒有 `dest_subdir` | 僅用 `default_path_template` 生成。          |
| 想修改分類名稱                              | 改 `key`。我設定了大分類用【】括住                                |

  <br>②分類修改：<br>
```
{
  "key": "【向井康二】世一",
  "exts": [".xls", ".xlsx", ".pdf"],
  "rename": "{YYYY}{MM}{DD}_世一證據_{orig}{ext}",
  "path_template": "E:/Finance/{YYYY}/{YYYYMM}/向井康二",
  "present_rule": { "mode": "any" }
}
```

| 行（鍵名） | 範例內容 | 功能與說明 |
|-------------|-----------|-------------|
| `"key":` | `"【向井康二】世一"` | 類別名稱，顯示於主程式按鈕與右側清單中。 |
| `"exts":` | `[ ".xls", ".xlsx", ".pdf" ]` | 限定允許搬入的檔案類型。 |
| `"rename":` | `"{YYYY}{MM}{DD}_世一證據_{orig}{ext}"` | 定義搬移後檔案的自動命名規則。 |
| `"path_template":` | `"E:/Finance/{YYYY}/{YYYYMM}/向井康二"` | 指定此分類的目標資料夾路徑模板。 |
| `"present_rule":` | `{ "mode": "any" }` | 控制右側清單樹的狀態顯示條件。 `"any"`≥1個就√，`"count_at_least", "n": 3` ≥3個才√|
|`"dest_subdir":`|`"子目錄一級/子目錄二級"`|可刪，用於全局路徑下的子目錄。文檔放在`default_path_template`/`dest_subdir`的位置|


## 7、滑鼠右鍵SendTo
-WIN+R→輸入：`shell:sendto`
<br>-右鍵→New→Shortcut→輸入kinder_classify.py的位置。
<br>**logo設置**
<br>①文件準備：Photoshop準備好256*256的png，google search：png to ico。upload上去，download `*.ico`下來.
<br>②設置：剛剛SendTo裡面的shortcut→屬性（Alt+左鍵雙擊）→Chang Icon→選剛剛 `*.ico`

# 邏輯思路
## Classify功能
**流程：** 
<br>①輸入：拖曳檔案到 Listbox 或從 SendTo 傳入；加入 `self.files: List[Path]`。
<br>②選擇：預設「自動選中列表第一個」，避免沒選就全搬。
<br>③歸檔邏輯是按照年月組織，因此設置了下拉菜單來設置操作的年月→生成包含年月的目標資料夾，和包含年月的文件名稱。
<br>④分類：點某類別按鈕 → 對每個選中的檔案執行 搬移 + 重命名。
<br>⑤回饋：狀態列顯示結果；右側 Checklist 重新統計。
<br>**拖拽和SendTo**
<br>①拖曳：`tkinterdnd2` 綁 `<<Drop>>`，用 `self.tk.splitlist(event.data)` 解析多檔路徑；去重後加入 `self.files`。
<br>②SendTo 單例：
<br>程式啟動時嘗試在 `127.0.0.1:port` 開TCP服務；若已被占用，表示已有打開的窗口 ⇒ 對該埠送 JSON（檔案列表），然後本程序退出。
<br>主實例（已開的窗口）收到 JSON 用 app.after(0, app._add_files, files) 追加到待辦清單。
<br>**操作棧（撤銷/重做）** 
<br>①移動和避免覆蓋：`move_with_conflict()` 在目標目錄存在重名就在名字後加`-1/-2...`
<br>②撤回/重做：用操作棧記錄每一步`{"orig": 原路徑, "dst": 目標, "current": 當前位置}`
<br>-撤回：把文件從「當前位置」移回`orig`，如過衝突則附加`-undo1`,然後回填到`代辦列表`
<br>-重做：把撤回的文件移到`dst`，同時從`代辦列表`移出。

## 文件清單Checklist
**統計：** 遍歷每類目標目錄，按照`前綴+擴展名`過濾計數，結合`present_rule`判定狀態✅/⬜
<br>**Treeview:** 
<br>-兩欄：左 文件類別（`#0`）、右 狀態[數量]（`values=("✅[5]",)` 或 `("⬜[0]",)`）。
<br>-群組：優先讀 `item["group"]`；無則從 `key` 中 `【】` 解析；每組插一個「—— 組名 ——」父節點。
<br>-互動：雙擊非組節點 → `os.startfile(target_dir(item))` 打開目錄。

