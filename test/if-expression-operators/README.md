# IF Expression Operators Test

## 測試目的
驗證 IF 條件運算支援：
- `!=`
- `>` `<` `>=` `<=`
- `and` `or` `not`
- 巢狀括號 `()`

## 測試檔案
- 輸入：`commands_if_expr_sample.xlsx`
- 輸出（歷史）：`output_if_expr_sample.xlsx`
- 輸出（建議最新）：`output_latest.xlsx`

## 執行方式
```bash
python3 redfish_agent.py -u <USER> -p <PASS> -r <ROOT_URL> \
  -f ./test/if-expression-operators/commands_if_expr_sample.xlsx \
  -o ./test/if-expression-operators/output_latest.xlsx
```

## 預期重點
- IF 條件可正確進入 PASS 或 FAIL 分支。
- MESSAGE 輸出可看到替換後條件值。
