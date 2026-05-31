# GUI Runner

This folder contains a desktop GUI test runner for `redfish_agent.py`.

## Features

- English UI
- Add multiple targets (`Root URL`, `Username`, `Password`)
- Connection status list (`Connection URL`, `Username`, `Connection Status`)
- Modify selected target from target list
- Remove selected target from target list
- Import targets from JSON
- Export targets to JSON
- Select one command Excel file
- Run tests in parallel for all targets
- Save each target output Excel into `gui/output/`
- Generate one final HTML report in `gui/output/`
- Open the latest generated HTML report directly from GUI

## Important Note About Batch

The GUI uses `redfish_agent.py` as-is.
If your test design requires one test item per batch, define your Excel commands accordingly using:

- `BATCH_START(<BatchName>)`
- `BATCH_END(<BatchName>)`
- `BATCH(<BatchName>)`

## Run

From repository root:

```bash
python3 gui/redfish_gui_runner.py
```

## Output

Generated files are placed in:

- `gui/output/output_<target>_<timestamp>.xlsx`
- `gui/output/report_<timestamp>.html`
