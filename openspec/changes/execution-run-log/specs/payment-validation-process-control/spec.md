# Delta for Payment Validation Process Control

## ADDED Requirements

### Requirement: Execution correlation columns

The process control workbook (`tblControlProcesosPagos`, row 2) MUST support two additive columns: `ExecutionId` and `ExecutionLogPath`. Setup idempotent MUST create or repair these columns without overwriting row 2 data. Existing operational semantics of `EstadoProceso`, `ProcessKey`, `ProcessId`, job ID columns, and manifest path MUST NOT change.

`ExecutionId` MUST store the UUID of the active validation execution for correlation across endpoints. `ExecutionLogPath` MUST store the relative SharePoint path to the consolidated JSON in `04 LOGS`.

#### Scenario: Setup repairs legacy workbook

- GIVEN a control workbook created before this change
- WHEN setup or repair runs
- THEN `ExecutionId` and `ExecutionLogPath` columns MUST exist in the table schema
- AND existing row 2 values for other columns MUST remain unchanged

#### Scenario: Generate persists correlation

- GIVEN `EXECUTION_RUN_LOG_ENABLED=true` and new execution accepted
- WHEN Generate reserves `execution_id` and initializes log
- THEN control row 2 MUST be updated with matching `ExecutionId` and `ExecutionLogPath` before or when Generate completes successfully
- AND downstream steps MUST read these fields to locate the log

#### Scenario: Flag disabled no column writes for logging

- GIVEN `EXECUTION_RUN_LOG_ENABLED=false`
- WHEN any step runs
- THEN MUST NOT write `ExecutionId` or `ExecutionLogPath` for logging purposes
- AND control behavior MUST match current system

---

### Requirement: Process control snapshot reads execution fields

`ProcessControlSnapshot` (or equivalent read model) MUST expose `execution_id` and `execution_log_path` when columns are present. When empty, values MUST be treated as missing—not inferred from bank or date.

#### Scenario: Snapshot exposes execution fields

- GIVEN control row 2 has `ExecutionId` and `ExecutionLogPath` populated
- WHEN `read_process_control_snapshot` runs for the bank
- THEN snapshot MUST include non-empty `execution_id` and `execution_log_path`

#### Scenario: Empty execution fields

- GIVEN legacy control with empty `ExecutionId`
- WHEN snapshot is read
- THEN `execution_id` MUST be empty
- AND consumers MUST NOT infer correlation from other fields alone

---

## MODIFIED Requirements

None. Operational state transitions, terminal states, idempotency keys, and `MergeManifestPath` semantics are unchanged.

## REMOVED Requirements

None.
