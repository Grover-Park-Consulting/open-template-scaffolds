---
id: dialog-blocks-automation-read-it-from-outside
mechanism: vba-generation
kind: guardrail
confidence: Observed
status: active
lesson: 'A MsgBox in code driven through automation blocks the call, and a dialog watchdog dismisses it before its text can be read from inside Access. Read it from outside the process: a background script polling for a window by title and dumping its controls'' text, launched before the triggering call. A background process launched with a trailing & did not survive the call returning.'
applies_when: A build drives Access through automation and code can raise a message box.
evidence: A watchdog dismissed the dialog before inside-Access reads could see it; the outside poller captured the text.
environment:
- Build driven through an Access MCP server or COM automation. Last checked 2026-09-22. Later builds reused the method and are not counted.
sources:
- error-logging-scaffold 09-22
confirmations: []
contradictions: []
---
