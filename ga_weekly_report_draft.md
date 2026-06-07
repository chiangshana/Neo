# 霓淨思工廠訂單排程｜本週 GA 實作版簡報草稿

這份草稿參考你們上次的簡報模板結構，改寫成這週要報告的「GA 演算法實作進度」版本。

建議頁數：8～10 頁  
如果你的報告時間只有 3～5 分鐘，優先講第 1、2、3、5、6、8、9、10 頁。

---

## Slide 1｜封面

### 標題
霓淨思 Neogence  
工廠訂單排程  
本週進度報告：遺傳算法（GA）實作

### 可放內容
- 第二組
- 報告人：`你的名字`
- 日期：`本週報告日期`

### 講者備註
這週的重點不是再介紹理論，而是說明我們已經把 GA 從上次的 pseudo code，推進到可以實際跑正式版、輸出排程表、並做驗證的版本。

---

## Slide 2｜回顧｜上次到本次的 GA 重點

### 標題
回顧｜上次到本次的 GA 重點

### 內容
上次進度：
- 以 pseudo code 形式整理 GA 流程
- 初步確認染色體可用「訂單排序」表示
- 開始整理限制式與懲罰項方向

本次完成：
- 將 GA 實作成可執行的排程器
- 建立資料讀取、建模、排程、驗證、輸出的一條龍流程
- 可輸出正式版排程表、每日負荷表、異常訂單表、欄位說明表

本次核心進展：
- 染色體：訂單排序
- 解碼器：依排序逐張訂單安排製造、排瓶、充填
- 適應度：總加班 + 不可行懲罰 + 目標偏離懲罰
- 驗證：確認結果是否真的符合核心限制

### 講者備註
這一頁的目的，是讓大家快速知道我們已經從「想法」進到「可以執行、可以檢查」的階段。

---

## Slide 3｜目前使用的方法｜遺傳算法（GA）建模架構

### 標題
目前使用的方法｜遺傳算法（GA）建模架構

### 內容
GA 在本專案中的表示方式：
- 染色體：一條訂單排序序列
- 個體：一個完整的訂單排序方案
- 解碼方式：依照排序，逐張訂單往時間軸中安插

解碼後會產生四類工序：
- `Cat1`：預攪
- `Cat2`：攪料清洗
- `Cat3`：排瓶
- `Cat4`：充填

模型評分邏輯：
- `Objective = total_overtime + infeasible_penalty + target_penalty`

其中：
- `total_overtime`：總加班分鐘數
- `infeasible_penalty`：違反硬限制的高額懲罰
- `target_penalty`：偏離目標加班上限的懲罰

### 建議視覺
- 一張簡單流程圖：
  `訂單排序 → 解碼器 → 產生排程 → 計算 overtime / penalty → fitness`

### 講者備註
我們沒有直接把染色體設計成「每張訂單的具體日期」，而是用「排序」加上解碼器來做。這樣比較容易重排，也比較符合 GA 的搜尋特性。

---

## Slide 4｜目前使用的方法｜GA 初始族群與演化設定

### 標題
目前使用的方法｜GA 初始族群與演化設定

### 內容
本次 GA 主要參數：
- `population_size = 40`
- `generations = 120`
- `mutation_rate = 0.15`
- `crossover_rate = 0.85`
- `seed = 42`
- `elite_count = 4`
- `tournament_size = 3`

初始族群來源：
- 依交期排序
- 依物料齊料點排序
- 依原料齊料點排序
- 部分負荷導向排序
- 隨機打散排序

這樣做的原因：
- 不讓所有初始解都太像
- 保留多樣性，避免太早收斂
- 同時保留一些具業務直覺的起始排序

### 可放程式片段
```python
population.append(sorted(orders, key=lambda item: (item.due_date, item.order_id)))
population.append(sorted(orders, key=lambda item: (item.material_ready_date, item.due_date)))
population.append(sorted(orders, key=lambda item: (item.raw_ready_date, item.due_date)))
```

```python
@dataclass
class GeneticScheduler:
    population_size: int = 40
    generations: int = 120
    mutation_rate: float = 0.15
    crossover_rate: float = 0.85
    elite_count: int = 4
    tournament_size: int = 3
    seed: int = 42
```

```python
def _initialize_population(self, orders: list[Order]) -> list[list[Order]]:
    base = sorted(orders, key=lambda item: (item.due_date, item.material_ready_date, item.order_id))
    population = [base]
    population.append(sorted(orders, key=lambda item: (item.material_ready_date, item.due_date)))
    population.append(sorted(orders, key=lambda item: (item.raw_ready_date, item.due_date)))
```

### 這頁若想放更多 code，建議這樣排
- 左半：GA 主要參數 + 初始族群設計理由
- 右半：放 `GeneticScheduler` 參數設定與 `_initialize_population()`
- 同一頁最多保留 2 段短 code，就不會太擠

### 講者備註
這邊想傳達的是，我們不是完全隨機亂排，而是把人工排程的直覺順序也納進初始族群，讓搜尋更穩。

---

## Slide 5｜目前使用的方法｜GA 解碼器與限制處理

### 標題
目前使用的方法｜GA 解碼器與限制處理

### 內容
目前解碼器已處理的核心限制：

資料與日期：
- 製造不得早於 `raw_ready_date`
- 排瓶/充填不得早於 `material_ready_date`
- 所有工序不得晚於 `due_date`

製造流程：
- 預攪一定要在攪料前一天
- 製造與充填人力分開，不共用
- 製造端依機器與人力容量安排

充填流程：
- 排瓶原則上要在充填前一天
- 若資料窗過緊，允許特定情況下同天排瓶＋充填
- 排瓶最多投入 2 人
- 充填與排瓶依 5 分鐘 slot 切分與安排

輸出邏輯：
- 對外報表統一改成整數顯示
- 同日、同訂單、同工序會先合併，避免輸出過碎
- reference format 會同時顯示「標準配置人數」與「實際安排人數」

### 建議視覺
- 放一張工序示意圖：
  `Cat1 預攪 → Cat2 攪料清洗 → Cat3 排瓶 → Cat4 充填`

### 講者備註
這一頁可以強調：GA 不是只做排序，它背後有一個 greedy decoder，把排序翻譯成真的排程結果。

---

## Slide 6｜目前使用的方法｜GA 適應度函數與驗證機制

### 標題
目前使用的方法｜GA 適應度函數與驗證機制

### 內容
適應度函數目前包含三層：

1. 總加班分鐘數
- 調製、排瓶、充填分開統計
- 加班越少越好

2. 不可行懲罰
- 若違反硬限制，直接給很大的 penalty
- 目標是先求可行，再談優化

3. 目標偏離懲罰
- 調製目標：2500 分鐘
- 充填目標：2585 分鐘
- 若超出目標，額外加懲罰

目前 validation 已檢查：
- 每張可排訂單是否真的有排進去
- 充填分鐘數是否與建模一致
- 排瓶工作量是否與建模一致
- 預攪是否真的在前一天
- 排瓶是否在充填前一天或允許的同天情境
- 是否提早於 ready date 或晚於 due date

### 建議視覺
- 左邊放 fitness 組成
- 右邊放 validation 清單

### 講者備註
這一頁要強調一個觀念：GA 找到低 fitness 的解，不代表一定合理，所以我們另外加了 validation，避免出現「看起來分數低，但其實違規」的結果。

---

## Slide 7｜本週 Debug 與實作修正

### 標題
本週 Debug 與實作修正

### 內容
本週已修正的重要問題：

1. 報表顯示問題
- 原本 reference format 曾把排瓶標準人數與實際人數混在一起
- 現在已拆成：
  - `標準配置人數`
  - `實際安排人數`

2. 同日排瓶/充填顯示順序
- 修正 reference format 的 slot 顯示邏輯
- 避免出現「表上看起來先充填、後排瓶」的誤讀

3. 總加班輸出口徑統一
- `schedule_summary.csv` 與主控台摘要都統一顯示為整數

4. 正式版執行資訊紀錄
- 新增 `run_metadata.csv`
- 可記錄開始時間、結束時間、總秒數、GA 參數

5. 輸出可讀性
- 每次輸出自動帶時間戳記資料夾
- 自動產生 `欄位備註表/`，方便組員與業主核對

### 講者備註
這一頁是實作感最強的一頁，可以讓老師看到我們不是只跑模型，也有實際 debug 與改善可解讀性。

---

## Slide 8｜目前結果｜正式版輸出摘要

### 標題
目前結果｜正式版輸出摘要

### 內容
最新正式版輸出：
- 輸出資料夾：`ga_output_formal_20260421_110003`

最新正式版結果：
- 原始訂單數量：`103`
- 實際參與 GA 的訂單數量：`100`
- 異常訂單筆數：`5`
- 排程驗證問題數：`0`
- 總加班分鐘數：`1774`
- 調製超時分鐘數：`0`
- 排瓶超時分鐘數：`1742`
- 充填超時分鐘數：`31`
- 週末加班分鐘數：`769`
- 正式版執行時間：`00:06:00`

這代表：
- 目前模型已能產生可行解
- 主要瓶頸仍然集中在排瓶，而不是調製
- 充填端本身的 overtime 已被壓到相對低

### 建議視覺
- 放一張黑底白字的摘要框，格式比照你們上次 LP 節錄頁的風格

### 講者備註
這頁是你最重要的結果頁，可以直接講「現在我們的 GA 版本已經能穩定輸出可行排程，且 validation 為 0」。

---

## Slide 9｜目前限制與風險

### 標題
目前限制與風險

### 內容
目前仍有幾個限制：

1. 清洗邏輯仍屬近似
- 目前有 `manufacturing_cleaning_audit.csv`
- 也有清洗 penalty
- 但尚未完全把清洗當作硬工序直接塞入時間表

2. 來源資料映射仍有歧義
- `order_source_links.csv` 仍顯示部分訂單不是唯一對應
- 代表某些建模參數仍有資料清理空間

3. 加班集中在排瓶端
- 目前排瓶是最主要瓶頸
- 後續可能要再優化人數配置與拆分策略

4. 無解訂單仍需分流
- 有些訂單是資料窗本身就不可行
- 這些會被 anomaly 標出，但不代表模型把它們解掉了

### 講者備註
這頁建議講得誠實一點。這不是模型失敗，而是告訴老師我們知道現在還有哪些近似與限制。

---

## Slide 10｜Next Step｜GA 下一步優化方向

### 標題
Next Step｜GA 下一步優化方向

### 內容
下一步預計優化：

1. 把清洗從 penalty / audit 升級成更明確的硬限制

2. 進一步改善排瓶瓶頸
- 更細的人力配置規則
- 同批次合併處理
- 更貼近現場的排瓶安排方式

3. 持續補強來源資料映射
- 降低產品對應歧義
- 提高建模參數可信度

4. 強化對外說明能力
- 報表格式持續貼近業主投影片
- 增加更容易檢查合理性的輸出表

### 講者備註
最後收在「目前 GA 已經從概念變成可跑、可驗證、可輸出的版本，下一步是把現場細節再逐步加回去」。

---

## 3 分鐘口頭版摘要

如果你只想用很短的方式講，可以用下面這段：

> 這週我主要報告的是 GA 演算法的實作進度。上次我們還停留在 pseudo code，這週已經把 GA 實作成可以正式跑的排程器。  
> 我們目前把染色體設計成訂單排序，再透過 decoder 依序安排預攪、攪料清洗、排瓶和充填。  
> fitness 主要由總加班、不可行懲罰與目標偏離懲罰組成，並另外加上 validation 去確認結果是否真的符合限制。  
> 目前正式版已經可以輸出可行排程，validation 是 0，總加班分鐘數是 1774 分鐘，瓶頸主要集中在排瓶。  
> 本週也修正了報表人數顯示、同日排瓶與充填順序，以及執行時間紀錄等問題。  
> 下一步會把清洗與更多現場細節再逐步加回模型中，讓 GA 結果更貼近實務排程。  

---

## 建議可直接截圖放進簡報的檔案

- `ga_output_formal_20260421_110003/schedule_summary.csv`
- `ga_output_formal_20260421_110003/best_schedule_reference_format.csv`
- `ga_output_formal_20260421_110003/schedule_validation_summary.csv`
- `ga_output_formal_20260421_110003/run_metadata.csv`
- `ga_output_formal_20260421_110003/欄位備註表/`
