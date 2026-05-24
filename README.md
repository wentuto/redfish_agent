# Redfish Agent

Redfish Agent executes command-driven Redfish workflows from an Excel file and writes execution results to an output Excel file.

## Key Features

- HTTP command execution: `GET`, `POST`, `PATCH`, `PUT`, `DELETE`
- Control flow commands: `IF`, `ELSE`, `ENDIF`
- Batch commands: `BATCH_START(<BatchName>)`, `BATCH_END(<BatchName>)`, `BATCH(<BatchName>)`
- Batch terminal commands: `SUCCESS`, `ERROR`
- Message command: `MESSAGE` (prints substituted payload into output `Response`)
- Context substitution: `CONTEXT`, `BATCH.CONTEXT`, `RESPONSE`, `STATUSCODE`
- Request assignment rules (multi-line) for context updates

## Installation

1. Install Python 3.8+.
2. Clone the repository.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

Dependencies:

- `requests`
- `openpyxl`

## Command Workbook Format

Input workbook columns (in order):

1. `Method`
2. `Endpoint`
3. `Payload`
4. `Request`

Header row example:

- `Method | Endpoint | Payload | Request`

## Macro and Context Overview

Supported macro namespaces:

- `${CONTEXT.<key>}`
- `${BATCH.CONTEXT.<key>}`
- `${RESPONSE.<json_path>}`
- `${STATUSCODE}`

Legacy `${username.id}` remains compatible.

## Control Commands Overview

- `IF`
    - Condition is written in `Payload`
    - Supports `==`, `!=`, `>`, `<`, `>=`, `<=`, `and`, `or`, `not`, and `()`
- `BATCH(<BatchName>)`
    - Argument is written in `Payload`
- `MESSAGE`
    - Uses substituted `Payload` and writes it to output `Response`

## Run

Default input/output:

```bash
python3 redfish_agent.py -u <user> -p <password> -r <root_url>
```

Custom input/output:

```bash
python3 redfish_agent.py -u <user> -p <password> -r <root_url> -f <input.xlsx> -o <output.xlsx>
```

Example:

```bash
python3 redfish_agent.py -u admin -p Kaori -r https://127.0.0.1:8000 -f ./test/if-batch-message-flow/commands_sample_if_batch.xlsx -o ./test/if-batch-message-flow/output_latest.xlsx
```

## Documentation

- Chinese spec: [docs/development_spec.md](docs/development_spec.md)
- English spec: [docs/development_spec_en.md](docs/development_spec_en.md)

## Test Cases

The repository uses one folder per test case under [test](test):

- [test/if-batch-message-flow](test/if-batch-message-flow)
- [test/if-expression-operators](test/if-expression-operators)
- [test/message-substitution](test/message-substitution)

Each test folder includes:

- command workbook
- output workbook(s)
- `README.md` for test purpose and run command

