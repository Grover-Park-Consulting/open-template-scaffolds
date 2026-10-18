---
id: lock-the-parent-row-before-rolling-up-children
mechanism: derived-values
kind: guardrail
confidence: Inferred
status: active
lesson: A stored total computed from child rows can come out one short, with no error, when two writers each compute it before seeing the other's committed row. Take a write lock on the parent row first (a no-change UPDATE inside the writer's transaction), then read the children, so writers queue and the second one reads the first one's committed row. A writer that has to wait gets the engine's normal lock wait; if that times out the error reaches the caller instead of a silent wrong total.
applies_when: A stored value on a parent row is recomputed from its child rows inside a per-write transaction, and more than one writer can act on the same parent at once.
evidence: The builder added the lock from reasoning about the missing-row case. The unlocked failure was never reproduced. The locked path was exercised only with a forced second connection that inserted a rival row mid-write; wait and timeout behavior were not tested beyond that.
environment: []
sources:
- stocktake-scan-outcome-first 10-04
confirmations: []
contradictions: []
---
