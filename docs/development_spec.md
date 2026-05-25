# Redfish Agent 開發規格文件

版本: v1.0-draft  
日期: 2026-05-24  
適用範圍: `redfish_agent.py` 與 Excel 指令執行流程

## 1. 文件目的

本文件定義 `execute_redfish()` 的擴充行為，讓系統在既有 HTTP method 之外，支援流程控制、批次執行、變數替換與上下文保存。

本文件作為開發實作的唯一依據，重點包含:

- Excel 欄位定義
- Method 語法與分類
- Substitution 與 Context 規則
- Request 欄位賦值規則
- IF/ELSE/ENDIF 控制流程
- BATCH 定義與執行流程
- SUCCESS/ERROR 終止語意
- Output 寫入規則
- 錯誤處理與限制

## 2. 術語定義

- Global Context (`CONTEXT`): 主流程共享變數空間。
- Batch Context (`BATCH.CONTEXT`): 單次 batch 執行時的區域變數空間。
- Substitution: 將字串中的 `${...}` MACRO 替換成實際值。
- HTTP Command: 會送出 HTTP request 的 method。
- Control Command: 不送 HTTP request 的流程控制 method。

## 3. Excel 格式定義

輸入 Excel 應包含以下欄位，順序固定如下:

1. `Method`
2. `Endpoint`
3. `Payload`
4. `Request`

說明:

- `Method`: HTTP method 或控制指令。
- `Endpoint`: HTTP endpoint，或控制指令的參數區。
- `Payload`: HTTP request body (JSON 字串或空值)。
- `Request`: 可多行的賦值規則，用於保存本次執行產生的資料。

## 4. Method 類型與語法

### 4.1 HTTP Methods

第一版至少支援:

- `GET`
- `POST`
- `PATCH`
- `PUT`
- `DELETE`

### 4.2 Control Methods

新增以下控制指令:

- `IF`
- `ELSE`
- `ENDIF`
- `BATCH_START(<BatchName>)`
- `BATCH_END(<BatchName>)`
- `BATCH(<BatchName>)`
- `SUCCESS`
- `ERROR`
- `MESSAGE`

控制指令參數欄位規則:

- `IF` 的 `<condition>` 寫在 `Payload` 欄位。
- `BATCH(<BatchName>)` 的 `<argument>` 寫在 `Payload` 欄位。
- `MESSAGE` 的輸出字串寫在 `Payload` 欄位。
- `BATCH_START/BATCH_END/ELSE/ENDIF/SUCCESS/ERROR` 不使用 `Payload`。

`MESSAGE` 語意:

- 當讀到 `MESSAGE` 時，不會發送 HTTP request。
- 會先對 `Payload` 做 substitution。
- substitution 後的結果會寫入 output Excel 的 `Response` 欄位。

## 5. 變數空間與 Macro 規格

### 5.1 支援的變數命名空間

- `${CONTEXT.<key>}`
- `${BATCH.CONTEXT.<key>}`
- `${RESPONSE.<json_path>}`
- `${STATUSCODE}`

### 5.2 Endpoint macro 規格調整

原先 `${username.id}` 改為可擴充命名空間形式，推薦使用:

- `${CONTEXT.<name>.id}`
- `${BATCH.CONTEXT.<name>.id}`

若仍需保留舊格式，應在相容層轉換為新格式後再解析。

### 5.3 Substitution 基本規則

- 所有可替換欄位在執行前都要先做 substitution。
- 可替換欄位包含: `Endpoint`, `Payload`, `Request`, `IF condition`, `BATCH argument`。
- substitution 發生錯誤 (變數不存在、路徑不存在) 時，應記錄錯誤並依錯誤策略中止或跳過。

## 6. Request 欄位規格

`Request` 欄位可包含多行，每行為一條賦值語句:

`${TARGET} = <VALUE>`

範例:

`${CONTEXT.code1} = ${STATUSCODE}`
`${CONTEXT.enable1} = ${RESPONSE.ServiceEnabled}`
`${BATCH.CONTEXT.device} = {"id": 1, "name": "A"}`

### 6.1 第一個關鍵規則 (執行順序)

每一行 Request 規則都必須依序執行以下流程:

1. 先解析右值 substitution。
2. 右值解析為最終值 (純值或 JSON)。
3. 寫入左值對應目標變數。

注意:

- 同一個 Request 欄位中的多行會依行號順序執行。
- 後續行可以讀取前面行剛寫入的值。

### 6.2 左值限制

左值只允許可寫入目標:

- `${CONTEXT.<key>}`
- `${BATCH.CONTEXT.<key>}`

不允許:

- `${STATUSCODE}` (唯讀)
- `${RESPONSE...}` (唯讀)

### 6.3 右值來源

右值可為:

- `${STATUSCODE}`
- `${RESPONSE.<json_path>}`
- `${CONTEXT.<key>}`
- `${BATCH.CONTEXT.<key>}`
- JSON literal
- 一般字串或數值 literal

## 7. IF / ELSE / ENDIF 規格

### 7.1 語法

`IF` 語法:

`Method = IF`
`Payload = <boolean_expression>`

範例:

`Method: IF`
`Payload: (${CONTEXT.code1} >= 200 and ${CONTEXT.code1} < 300) and not (${CONTEXT.enable1} == false)`

### 7.2 第二個關鍵規則 (條件運算限制)

IF 條件支援以下比較運算:

- `==`
- `!=`
- `>` `<` `>=` `<=`

也支援布林運算與巢狀表達式:

- `and`
- `or`
- `not`
- `()`

範例:

- `(${CONTEXT.code1} >= 200 and ${CONTEXT.code1} < 300) and not (${CONTEXT.enable1} == false)`

目前不支援:

- 算術運算 (`+`, `-`, `*`, `/`)
- 函式呼叫
- 自訂識別字直接存取 (需透過 `${...}`)

### 7.3 控制流語意

- IF 成立: 執行 IF 區塊，跳過 ELSE 區塊。
- IF 不成立: 跳過 IF 區塊，若有 ELSE，執行 ELSE 區塊。
- ENDIF: 結束當前 IF 區塊。

### 7.4 配對規則

- `ELSE` 必須對應最近且未關閉的 `IF`。
- `ENDIF` 必須對應最近且未關閉的 `IF`。
- 若配對失敗，視為語法錯誤。

## 8. BATCH 規格

### 8.1 BATCH 定義

`BATCH_START(BatchName)` 與 `BATCH_END(BatchName)` 之間所有列形成一個可重用 batch 定義。

### 8.2 BATCH 收集階段

當主流程讀到 `BATCH_START(BatchName)`:

- 進入定義收集模式。
- 將直到對應 `BATCH_END(BatchName)` 之間的列存入 batch registry。
- 收集階段不執行其中 HTTP 或控制指令。

### 8.3 BATCH 執行

`BATCH(BatchName)` 觸發執行，且 `argument` 由 `Payload` 欄位提供:

1. 查找 batch registry 內的 `BatchName`。
2. 對 `argument` 做 substitution。
3. 初始化 `${BATCH.CONTEXT}`。
4. 將 `argument` 寫入 `${BATCH.CONTEXT}` (建議鍵名 `input` 或展開為物件)。
5. 依序執行 batch 內部指令。
6. 在遇到 `SUCCESS` 或 `ERROR` 時結束 batch。

### 8.4 BATCH Context 行為

- 每次 `BATCH(...)` 呼叫都建立新的 `${BATCH.CONTEXT}` 實例。
- batch 結束後 `${BATCH.CONTEXT}` 不應污染下一次 batch 執行。
- batch 內 `Request` 預設建議寫入 `${BATCH.CONTEXT...}`。

`BATCH` 列範例:

- `Method: BATCH(AccountAudit)`
- `Payload: {"target_user":"admin","from_status":${CONTEXT.last_status}}`

### 8.5 SUCCESS / ERROR

- `SUCCESS`: 表示當前 batch 成功完成，立即結束 batch。
- `ERROR`: 表示當前 batch 失敗，立即結束 batch。
- 執行到 `SUCCESS` 或 `ERROR` 時，需將 `${BATCH.CONTEXT}` 序列化後寫入 output `Response` 欄位。

## 9. HTTP 執行流程

當 `Method` 為 HTTP method 時，流程如下:

1. 對 `Endpoint`, `Payload`, `Request` 做 substitution。
2. 驗證與解析 `Payload` JSON。
3. 發送 HTTP request。
4. 接收 `status_code` 與 `response_body`。
5. 生成 `${STATUSCODE}` 與 `${RESPONSE...}` 可讀內容。
6. 執行 `Request` 欄位規則更新 context。
7. 寫入 output row。

## 10. Output Excel 規格

輸出檔建議欄位如下:

1. `Method`
2. `Endpoint`
3. `Payload`
4. `Request`
5. `Status Code`
6. `Response`

若為 batch 結束事件 (`SUCCESS` / `ERROR`):

- `Method` 填入實際執行的 `BATCH(<BatchName>)`
- `Response` 填入 `${BATCH.CONTEXT}` JSON 字串

若主流程執行 `BATCH(<BatchName>)`，output Excel 應額外寫入 batch 範圍標記列:

- 進入 batch 前寫入一列 `BATCH_START(<BatchName>)`
- batch 結束後寫入一列 `BATCH_END(<BatchName>)`

這兩列用來界定 batch 在 output 中的實際執行範圍。

若 `Method = MESSAGE`:

- `Status Code` 欄位填入 `MESSAGE`
- `Response` 欄位填入 substitution 後的 `Payload` 字串

## 11. 錯誤處理規格

### 11.1 語法錯誤

以下視為語法錯誤:

- 無法解析的控制指令格式
- IF/ELSE/ENDIF 配對錯誤
- BATCH_START/BATCH_END 配對錯誤
- Request 左值不合法

處理建議:

- 記錄錯誤到 output
- 標示 `Status Code = Error`
- 依設定選擇繼續或中止

### 11.2 substitution 錯誤

- 變數不存在
- JSON path 不存在
- 型別不匹配

處理建議:

- 回填明確錯誤訊息
- 避免程式崩潰

### 11.3 HTTP 錯誤

- request exception
- timeout
- 非 JSON response

處理建議:

- 保留原始文字 response
- 仍寫 output

## 12. 相容性與遷移建議

- 舊格式 `${username.id}` 建議遷移到 `${CONTEXT.username.id}`。
- 若需要平滑遷移，可先保留舊格式解析，再輸出 deprecation warning。
- 舊格式 `IF <condition>` (把條件寫在 Method) 仍可相容，但建議遷移為 `Method=IF` 並將條件放到 `Payload`。
- 舊格式 `BATCH(BatchName, argument)` 仍可相容，但建議遷移為 `Method=BATCH(BatchName)` 並將 argument 放到 `Payload`。

## 13. 非目標 (v1 不支援)

- IF 複合條件 (`and`, `or`)
- 巢狀 BATCH 定義
- 以 expression language 執行任意運算
- 跨檔案 include

## 14. 測試案例建議

最少應覆蓋:

1. Request 多行賦值與前後依賴。
2. IF 成立/不成立 + ELSE 分支。
3. BATCH 定義與多次執行。
4. BATCH argument substitution。
5. SUCCESS / ERROR 正確終止與輸出。
6. `${RESPONSE.path}` 深層路徑取值。
7. substitution 失敗與錯誤輸出。

## 15. 實作優先順序

1. Request 解析器與 substitution 引擎。
2. IF/ELSE/ENDIF 流程控制。
3. BATCH 收集與執行器。
4. SUCCESS/ERROR 與 output 對齊。
5. 舊 macro 相容與遷移告警。
