# Execution Run Log Specification

## Purpose

Consolidated JSON audit trail in SharePoint `04 LOGS`: one `execution_id` = one file = one full validation flow (Generate through Apply). Fully additive behind `EXECUTION_RUN_LOG_ENABLED`; financial flow MUST NOT depend on logging success.

## Non-Functional Requirements

| ID | Requirement |
|----|-------------|
| NFR-1 | When `EXECUTION_RUN_LOG_ENABLED=false`, behavior MUST match current system (financial results, control states, manifest paths, idempotency, HTTP contracts). |
| NFR-2 | Logging failures MUST NOT fail business use cases; MUST retry write; MAY set `execution_log_status=WRITE_FAILED` and `possible_gaps=true`. |
| NFR-3 | Optimistic concurrency: GET JSON + eTag → modify → PUT `If-Match`; on 412 MUST re-read, merge, retry up to 5 times with short backoff. eTag MUST NOT be persisted in JSON. |
| NFR-4 | MUST NOT serialize secrets, tokens, base64 payloads, or full email bodies. `technical_message` max 4000 chars. |
| NFR-5 | Power Automate MUST remain a consumer only; no new request fields or PA responsibilities. |
| NFR-6 | Canonical manifest path `merge_manifest_{bank}_{date}.json` MUST NOT change. |

## Identities

| Identity | Scope | Rules |
|----------|-------|-------|
| `execution_id` | Full flow | Stable across all steps and user retries of same run; one file per execution. |
| `invocation_id` | Single endpoint call | New UUID per accepted HTTP invocation of a principal endpoint. |
| `job_id` | Async job | One per invocation; used for polling repair. |
| `attempt` | Step within execution | Assigned when invocation context is created; same value for STARTED and terminal of that invocation. |
| `process_id` | Business process | Separate from `execution_id`; MAY be null until Generate succeeds. |

## Functional Requirements

### Requirement: Execution file location and naming

The system MUST store exactly one JSON file per `execution_id` at:

`04 LOGS/execution_log_{bank_code}_{run_timestamp_utc}_{execution_id_short}.json`

The system MUST NOT create per-run subfolders, per-event files, ZIP, or JSONL.

#### Scenario: Two runs same bank same calendar day

- GIVEN `EXECUTION_RUN_LOG_ENABLED=true` and two completed Generate flows for `banco_bogota` on the same day
- WHEN both flows finish Apply or terminal state
- THEN two distinct files exist with different `execution_id_short` values
- AND neither file overwrites the other

#### Scenario: Two banks concurrent

- GIVEN simultaneous Generate for `banco_bogota` and `banco_bancolombia`
- WHEN both run with flag enabled
- THEN each bank MUST have independent `execution_id` and log file
- AND locks MUST be per bank

---

### Requirement: New execution creation

A new `execution_id` and log file MUST be created only when:

1. a valid Generate request is accepted;
2. no active execution exists for that bank;
3. bank lock is acquired;
4. control confirms a new run MAY start.

#### Scenario: New Generate creates log

- GIVEN flag enabled, bank idle, no active process
- WHEN `POST /generate/queue` succeeds
- THEN a new `execution_id` is reserved before the background task runs
- AND a new JSON file is initialized with `REQUEST_RECEIVED`, `QUEUED`, and later Generate events
- AND control receives `ExecutionId` and `ExecutionLogPath` before or with successful Generate completion

#### Scenario: Active process rejects second Generate

- GIVEN an active non-terminal execution for the bank
- WHEN a second Generate is requested
- THEN the system MUST behave as today (`active_process_exists`)
- AND MUST NOT create a new `execution_id` or log file
- AND MAY append `REQUEST_REJECTED` to the active log when correlation is safe

#### Scenario: Idempotent Generate reuses log

- GIVEN control has reusable revision for same `process_key` with `ExecutionId` and `ExecutionLogPath`
- WHEN Generate returns `already_generated`
- THEN MUST NOT create a new file or orphan log
- AND MUST append `SKIPPED_IDEMPOTENT` or `ALREADY_COMPLETED` to the existing log
- AND MUST NOT replace the existing file

#### Scenario: New flow after prior completion

- GIVEN prior execution reached terminal success or allowed new Generate
- WHEN a new valid Generate starts
- THEN MUST create new `execution_id` and new JSON file even if same bank and day

---

### Requirement: Double Generate protection

Sequence MUST be: acquire bank lock → read control → decide active execution → reserve `execution_id` → persist `ExecutionId`/`ExecutionLogPath` → initialize log → enqueue Generate → release lock.

#### Scenario: Simultaneous Generate calls

- GIVEN two concurrent Generate requests for the same bank
- WHEN both arrive while flag enabled
- THEN exactly one MUST create execution and log
- AND the other MUST receive current active-process behavior
- AND MUST NOT produce two active `execution_id` values for one bank

---

### Requirement: Automatic endpoint correlation

Finalize, Notify, Merge, Dry-run, and Apply MUST read `ExecutionId` and `ExecutionLogPath` from bank control after resolving `bank_code`. MUST NOT correlate by bank alone, `process_key`, file name, or `job_id`.

#### Scenario: Downstream step correlates via control

- GIVEN Generate completed and control updated
- WHEN Merge runs without PA sending `execution_id`
- THEN Merge events MUST append to the same JSON file referenced in control

#### Scenario: Legacy control without ExecutionId

- GIVEN control row lacks `ExecutionId`
- WHEN any downstream step runs
- THEN financial flow MUST continue unchanged
- AND logging MUST skip or warn without inventing `execution_id`
- AND MUST NOT attach events to another run

---

### Requirement: Global summary states

`summary.overall_status` MUST use: `RUNNING`, `WAITING_FOR_NEXT_STEP`, `SUCCEEDED`, `FAILED`, `BLOCKED`, `PARTIAL`, `COMPLETED_WITH_WARNINGS`. MUST NOT use `SUCCEEDED` for intermediate stages only.

`WAITING_FOR_NEXT_STEP` MUST include `waiting_for` and `next_expected_step`.

| After step | `waiting_for` | `next_expected_step` |
|------------|---------------|----------------------|
| Generate success | `SECRETARY_REVIEW` | `FINALIZE` |
| Finalize success | `ACCOUNTING_DOCUMENT_UPLOAD` | `NOTIFY` |
| Merge success | `DRY_RUN_EXECUTION` | `DRY_RUN` |
| Apply success | — | — (`overall_status=SUCCEEDED`) |

#### Scenario: WAITING after Generate

- GIVEN Generate succeeded
- WHEN summary is recomputed
- THEN `overall_status` MUST be `WAITING_FOR_NEXT_STEP`
- AND `waiting_for` MUST be `SECRETARY_REVIEW`
- AND `next_expected_step` MUST be `FINALIZE`

#### Scenario: SUCCEEDED only after Apply

- GIVEN flow completed through Merge only
- WHEN summary is recomputed
- THEN `overall_status` MUST NOT be `SUCCEEDED`
- AND WHEN Apply succeeds or `ALREADY_COMPLETED` for closed run THEN `overall_status` MUST be `SUCCEEDED` and `flow_completed_at` MUST be set

---

### Requirement: Timestamps

Root `dates` MUST include: `run_started_at`, `last_activity_at`, `last_terminal_event_at`, `flow_completed_at`, `process_date`, `report_date`, `bank_transaction_dates`. `flow_completed_at` MUST be set only on Apply success, Apply `ALREADY_COMPLETED` for closed execution, or explicit closure—not on recoverable failures.

#### Scenario: Recoverable Finalize failure

- GIVEN Finalize attempt 1 fails with `amount_mismatch`
- WHEN summary updates
- THEN `flow_completed_at` MUST remain null
- AND `last_terminal_event_at` MUST reflect the failed attempt

---

### Requirement: Event lifecycle per principal endpoint

Each principal endpoint MUST record when applicable: `REQUEST_RECEIVED`, `QUEUED`, `STARTED`, terminal status (`SUCCEEDED`, `FAILED`, `BLOCKED`, `PARTIAL`, `COMPLETED_WITH_WARNINGS`, `SKIPPED_IDEMPOTENT`, `ALREADY_COMPLETED`, `REQUEST_REJECTED`). Long steps MAY emit `CHECKPOINT` events.

#### Scenario: Merge retry preserves attempts

- GIVEN `execution_id=RUN-001`, Merge attempt 1 PARTIAL with `job_id=JOB-A`
- WHEN user retries Merge (attempt 2 succeeds with `job_id=JOB-B`)
- THEN both attempts MUST remain in `events[]` with `attempt=1` and `attempt=2` respectively
- AND both share the same `execution_id`

#### Scenario: Finalize fail then succeed

- GIVEN Finalize attempt 1 FAILED
- WHEN Finalize attempt 2 SUCCEEDED
- THEN log MUST show both with distinct `invocation_id` and `job_id`
- AND `last_successful_step` MUST be `FINALIZE` after attempt 2

---

### Requirement: Event idempotency

`event_id` MUST equal concatenation of: `execution_id`, `invocation_id`, `step`, `substep` (or `main`), `attempt`, `status`. Before append, if `event_id` exists, MUST NOT duplicate.

#### Scenario: Duplicate hook invocation

- GIVEN terminal event already logged for a `job_id`
- WHEN the same hook runs again
- THEN `events[]` length MUST not increase for that `event_id`

---

### Requirement: Revision and logging health

Each successful write MUST increment `logging.revision`. On write failures, MUST update `logging_health` with `possible_gaps`, `write_failures_count`, and status `PARTIAL` when applicable.

#### Scenario: Persistent log write failure

- GIVEN Graph write fails after retries
- WHEN business step completes successfully
- THEN business outcome MUST be unchanged
- AND job MUST expose `execution_log_status=WRITE_FAILED` at job root
- AND log JSON MUST indicate incomplete traceability when partially written

---

### Requirement: Server restart visibility

Incomplete invocations MUST remain visible: STARTED and CHECKPOINT without terminal event.

#### Scenario: Apply interrupted by restart

- GIVEN Apply attempt 1 STARTED and CHECKPOINT recorded, no terminal event
- WHEN server restarts and user retries Apply
- THEN attempt 1 events MUST remain
- AND attempt 2 MUST use new `invocation_id`, `job_id`, incremented `attempt`
- AND business idempotency MUST adopt prior work without deleting attempt 1

---

### Requirement: Job polling observability

`GET /jobs/{job_id}` MUST record polling metadata in the execution log when flag enabled: first poll, last poll, poll count, observed status transitions. MUST NOT write on every identical poll; MUST write on first poll, status change, terminal response, HTTP error, or every 5th identical poll.

#### Scenario: Polling records status transitions

- GIVEN PA polls a Merge job from queued through completed
- WHEN flag enabled
- THEN `job_polling[]` MUST capture observed statuses with counts and timestamps
- AND MUST NOT require PA changes

#### Scenario: Polling repairs missing terminal event

- GIVEN business job completed but terminal hook write failed
- WHEN `GET /jobs/{job_id}` observes terminal status
- THEN system MUST call repair logic to append missing terminal event if absent for that `job_id`

---

### Requirement: Log reconstruction

When control has `ExecutionId` and `ExecutionLogPath` but file is missing, system MUST attempt minimal reconstruction with `reconstructed=true`, `possible_gaps=true`, identity, control snapshot, known paths, and job IDs—MUST NOT fabricate unknown historical events.

#### Scenario: Deleted log file

- GIVEN control points to missing JSON path
- WHEN downstream step runs with flag enabled
- THEN MUST create minimal reconstructed log or append with gap markers
- AND MUST NOT invent prior step events

---

### Requirement: Artifacts and size limits

`artifacts` MUST be a list with latest state per `artifact_id`. Detailed lists capped at 100 entries with `truncated` metadata when exceeded. `events[]` MUST NOT be truncated.

#### Scenario: Large Apply artifact set

- GIVEN more than 100 amortization tables touched
- WHEN artifacts recorded
- THEN `artifacts` MUST include `total_count`, `included_count`, `truncated=true`
- AND events MUST still record per-table actions where emitted

---

### Requirement: Manifest snapshot

Canonical manifest MUST remain unchanged. Execution log MUST store `manifest_snapshot` summary (status, group counts, outputs, incomplete groups, dry-run eligibility) on Merge terminal events.

#### Scenario: Canonical manifest unchanged

- GIVEN Merge completes with flag enabled
- WHEN manifest uploaded
- THEN path MUST remain `merge_manifest_{bank_code}_{report_date}.json`
- AND snapshot MUST exist inside execution log

---

### Requirement: Optional job root fields

When flag enabled, job state at root MUST MAY include `execution_id`, `execution_log_path`, `execution_log_status`, `execution_log_warning` even when `result` is null. Initial `202` responses MUST NOT change.

#### Scenario: Failed job with null result

- GIVEN Generate job fails before result payload
- WHEN job fetched via GET
- THEN root MAY still expose `execution_id` and `execution_log_status` when known

---

### Requirement: Central execution step wrapper

Implementation SHOULD use a single wrapper (e.g. `execution_step`) that records STARTED, assigns attempt from context, handles exceptions as FAILED, supports PARTIAL/BLOCKED/COMPLETED_WITH_WARNINGS, emits CHECKPOINTs, and never suppresses original exceptions.

#### Scenario: Wrapper preserves business exception

- GIVEN use case raises `ValueError` inside wrapper
- WHEN flag enabled
- THEN FAILED event MUST be logged
- AND the same exception MUST propagate to job failure handling as today

---

### Requirement: Dry-run recording

Independent Dry-run endpoint: `step=DRY_RUN`, `substep=null`. Internal Apply dry-run: `step=APPLY`, `substep=DRY_RUN`.

#### Scenario: Dry-run blocked

- GIVEN dry-run gate blocks Apply
- WHEN terminal recorded
- THEN status MUST be `BLOCKED` or equivalent
- AND MUST indicate validation failed before writes

---

### Requirement: HTTP errors without safe correlation

`422` before router MUST log to Render with `request_id`; MUST NOT create `execution_id` without safe correlation. `409` MAY log to active execution log when correlation exists. `404` job MUST log to Render; MAY warn execution log if metadata allows recovery.

#### Scenario: Feature flag disabled

- GIVEN `EXECUTION_RUN_LOG_ENABLED=false`
- WHEN any endpoint runs
- THEN no JSON files MUST be created or updated
- AND all financial and HTTP behavior MUST match pre-feature system

---

## Mandatory acceptance scenarios (index)

Scenarios above cover: (1) new log, (2) concurrent Generate single run, (3) active process, (4) idempotent Generate, (5) new flow after completion, (6) Finalize retry, (7) Notify already sent, (8) Merge partial then success, (9) Dry-run blocked, (10) Apply partial/retry, (11) restart, (12) Apply adopt, (13) dedup, (14) eTag recovery (NFR-3), (15) log failure, (16–17) polling, (18) legacy control, (19) manifest canonical, (20) flag false, (21) two banks, (22) multiple daily files, (23) reconstruction, (24) sanitization (NFR-4), (25) PA optional fields, (26–28) WAITING/SUCCEEDED/retries.

## Schema reference (v1)

See proposal §5 for root fields: `schema_version`, identities, `bank`, `dates`, `summary` (+ `waiting_for`, `next_expected_step`), `events[]`, `artifacts[]`, `affected_entities[]`, `metrics`, `manifest_snapshot`, `job_polling[]`, `logging`, `logging_health`, optional `reconstructed`.

Event object MUST include: `event_id`, `sequence`, `step`, `substep`, `attempt`, `status`, `invocation_id`, `job_id`, timestamps, `duration_ms`, `severity`, optional `error`, `metrics`, `artifacts`, `affected_entities`, optional `checkpoint_label`.

## Open decisions

| ID | Topic | Default for design |
|----|-------|-------------------|
| OD-1 | Bank lock mechanism (extend JobManager vs dedicated lock service) | Per-bank asyncio lock + control read |
| OD-2 | `ABANDONED` inference offline only | No runtime ABANDONED in v1 |
| OD-3 | `power_automate_run_id` source | Optional header `X-PA-Run-Id` future; null in v1 |
| OD-4 | Global 422 handler scope | All payment-validation routers |

## Verdict

**Ready for `/sdd-design`** — requirements and scenarios are sufficient to design `execution_run_log.py`, control column migration, wrapper, and polling hooks without implementing code in this phase.
