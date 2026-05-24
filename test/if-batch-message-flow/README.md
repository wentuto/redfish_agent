# IF + BATCH + MESSAGE Flow Test

## 測試目的
驗證複合流程能力：
- IF/ELSE/ENDIF 分支控制
- BATCH 定義與執行
- BATCH 輸出順序（BATCH -> BATCH_START -> 內部 Method -> BATCH_END）
- MESSAGE 輸出與 substitution

## 測試檔案
- 輸入：`commands_sample_if_batch.xlsx`
- 輸出（歷史）：`output_sample_if_batch.xlsx`、`output_commands_sample_if_batch.xlsx`
- 輸出（建議最新）：`output_latest.xlsx`

## 執行方式
```bash
python3 redfish_agent.py -u <USER> -p <PASS> -r <ROOT_URL> \
  -f ./test/if-batch-message-flow/commands_sample_if_batch.xlsx \
  -o ./test/if-batch-message-flow/output_latest.xlsx
```

## 預期重點
- Output 中可看到 `BATCH(<BatchName>)` 摘要列。
- Output 中可看到 `BATCH_START(<BatchName>)` 與 `BATCH_END(<BatchName>)` 範圍列。
- `MESSAGE` 的 Response 欄位為 substitution 後內容。
