# Lesson lookups

**Who reads this:** the AI assistant, whose sync script builds the lookup tables of the lessons database from these rows; the developer, who approves changes to them.

A **lookup** is a short fixed list that a lesson field chooses from. Each section below names one table of the lessons database; its rows keep the numbers shown, because lessons and the forms refer to the numbers. The sync adds a missing row and never changes or removes one. A lesson file names a value by its text.

## tlkpLessonKind

| ID | Value |
|---|---|
| 1 | consequence |
| 2 | guardrail |

## tlkpLessonConfidence

| ID | Value |
|---|---|
| 1 | Inferred |
| 2 | Observed |
| 3 | Assumed |
| 4 | Potential |

## tlkpLessonStatus

| ID | Value |
|---|---|
| 1 | active |
| 2 | inactive |

## tlkpLessonOrigin

| ID | Value |
|---|---|
| 1 | Library |
| 2 | Own |
| 3 | Inspection |

## tlkpLessonApproval

| ID | Value |
|---|---|
| 1 | Approved |
| 2 | Proposed |
| 3 | Rejected |

## tlkpLessonEventKind

| ID | Value |
|---|---|
| 1 | Confirmation |
| 2 | Contradiction |
