# Redfish Agent Development Specification

Version: v1.0-draft  
Date: 2026-05-24  
Scope: `redfish_agent.py` command execution behavior

## 1. Purpose

This document defines the extended behavior of `execute_redfish()`.
In addition to standard HTTP methods, the executor supports control flow, batch execution, variable substitution, and context persistence.

## 2. Terminology

- Global Context (`CONTEXT`): Shared variable space across the main flow.
- Batch Context (`BATCH.CONTEXT`): Per-run variable space used only during one batch execution.
- Substitution: Replacing `${...}` macros with runtime values.
- HTTP Command: A method that sends an HTTP request.
- Control Command: A method that controls flow without sending an HTTP request.

## 3. Excel Input Format

The input workbook must use this column order:

1. `Method`
2. `Endpoint`
3. `Payload`
4. `Request`

Column meaning:

- `Method`: HTTP method or control method.
- `Endpoint`: HTTP endpoint, or optional parameter field for certain control commands.
- `Payload`: HTTP body (JSON string), or parameter field for certain control commands.
- `Request`: Multi-line assignment rules for storing values into context.

## 4. Supported Methods

### 4.1 HTTP Methods

- `GET`
- `POST`
- `PATCH`
- `PUT`
- `DELETE`

### 4.2 Control Methods

- `IF`
- `ELSE`
- `ENDIF`
- `BATCH_START(<BatchName>)`
- `BATCH_END(<BatchName>)`
- `BATCH(<BatchName>)`
- `SUCCESS`
- `ERROR`
- `MESSAGE`

Control method argument placement:

- `IF` condition is in `Payload`.
- `BATCH(<BatchName>)` argument is in `Payload`.
- `MESSAGE` text is in `Payload`.
- `BATCH_START/BATCH_END/ELSE/ENDIF/SUCCESS/ERROR` do not require `Payload`.

`MESSAGE` behavior:

- No HTTP request is sent.
- `Payload` is substituted first.
- The substituted result is written to output `Response`.

## 5. Variable Namespaces and Macros

### 5.1 Supported namespaces

- `${CONTEXT.<key>}`
- `${BATCH.CONTEXT.<key>}`
- `${RESPONSE.<json_path>}`
- `${STATUSCODE}`

### 5.2 Endpoint macro migration

Preferred endpoint macros:

- `${CONTEXT.<name>.id}`
- `${BATCH.CONTEXT.<name>.id}`

Legacy `${username.id}` remains backward compatible.

### 5.3 Substitution rule

Substitution is applied before execution for all supported fields:

- `Endpoint`
- `Payload`
- `Request` right-hand side values
- `IF` condition
- `BATCH` argument

## 6. Request Column Rules

Each non-empty line in `Request` is an assignment:

`${TARGET} = <VALUE>`

Examples:

- `${CONTEXT.code1} = ${STATUSCODE}`
- `${CONTEXT.enable1} = ${RESPONSE.ServiceEnabled}`
- `${BATCH.CONTEXT.device} = {"id": 1, "name": "A"}`

Execution order per line:

1. Substitute right-hand side.
2. Parse right-hand side as JSON literal or scalar.
3. Assign to left-hand target.

Writable targets:

- `${CONTEXT.<key>}`
- `${BATCH.CONTEXT.<key>}`

Read-only sources:

- `${STATUSCODE}`
- `${RESPONSE...}`

## 7. IF / ELSE / ENDIF

### 7.1 Syntax

- `Method = IF`
- `Payload = <boolean_expression>`

Example:

- `Payload: (${CONTEXT.code} >= 200 and ${CONTEXT.code} < 300) and not (${CONTEXT.enabled} == false)`

### 7.2 Supported expression operators

Comparison operators:

- `==`
- `!=`
- `>`
- `<`
- `>=`
- `<=`

Boolean operators:

- `and`
- `or`
- `not`

Grouping:

- Parentheses `()`

Not supported:

- Arithmetic operators (`+`, `-`, `*`, `/`)
- Function calls
- Direct custom identifiers (must use `${...}`)

### 7.3 Flow semantics

- IF true: execute IF block, skip ELSE block.
- IF false: skip IF block, execute ELSE block if present.
- ENDIF: closes current IF block.

## 8. BATCH

### 8.1 Definition

Commands between `BATCH_START(BatchName)` and `BATCH_END(BatchName)` are stored as a reusable batch definition.

### 8.2 Execution

`BATCH(BatchName)` executes that stored definition.

At execution time:

1. Resolve and parse `Payload` as batch argument.
2. Create new `${BATCH.CONTEXT}`.
3. Store argument into `${BATCH.CONTEXT}` (`input` key and expanded keys for objects).
4. Execute internal commands.
5. Stop at `SUCCESS` or `ERROR`.

### 8.3 SUCCESS / ERROR

Inside a batch only:

- `SUCCESS` ends batch with success status.
- `ERROR` ends batch with error status.

At batch end, `${BATCH.CONTEXT}` is serialized into output `Response`.

## 9. Output Workbook Format

Output columns:

1. `Method`
2. `Endpoint`
3. `Payload`
4. `Request`
5. `Status Code`
6. `Response`

For batch execution, output order is:

1. `BATCH(<BatchName>)` summary row (`Status Code` = `SUCCESS` or `ERROR`)
2. `BATCH_START(<BatchName>)`
3. Internal executed rows
4. `BATCH_END(<BatchName>)`

For `MESSAGE`:

- `Status Code` is `MESSAGE`
- `Response` is the substituted payload text

## 10. Error Handling

- Syntax errors (bad control format, IF/BATCH pairing issues, invalid Request target) are written to output with `Status Code = Error`.
- Substitution errors (missing variable/path/type mismatch) are written to output.
- HTTP errors (exceptions, timeout, non-JSON response) are written to output.

## 11. Backward Compatibility

Supported legacy forms:

- `${username.id}` macro
- `IF <condition>` inline in `Method`
- `BATCH(BatchName, argument)` inline in `Method`

Recommended migration:

- Use `Method = IF` with condition in `Payload`
- Use `Method = BATCH(BatchName)` with argument in `Payload`
