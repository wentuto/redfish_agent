# MESSAGE Substitution Test

## 測試目的
驗證 MESSAGE Method：
- 不發送 HTTP request
- Payload 先 substitution
- substitution 結果寫入 output 的 Response 欄位

## 測試檔案
- 輸入：`commands_message_sample.xlsx`
- 輸出（歷史）：`output_message_sample.xlsx`
- 輸出（建議最新）：`output_latest.xlsx`

## 執行方式
```bash
python3 redfish_agent.py -u <USER> -p <PASS> -r <ROOT_URL> \
  -f ./test/message-substitution/commands_message_sample.xlsx \
  -o ./test/message-substitution/output_latest.xlsx
```

## 預期重點
- Output 中 `Status Code` 為 `MESSAGE`。
- `Response` 為 substitution 後的字串。
