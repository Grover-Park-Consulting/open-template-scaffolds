Build scan-driven stocktake counting into my Access database using the outcome-first method.

Outcome: every scan is recorded and nothing scanned is silently dropped, duplicate scans of the same
item are recognized and not double-counted, a barcode that matches nothing known is held for review
rather than guessed at or discarded, and any product whose count strays outside its tolerance —
including one nobody scanned at all — is flagged for review.

Resources (library root): templates/stocktakescan/stocktake-scan-outcome-first.md (the template),
templates/stocktakescan/stocktake-schema.md (paired tables), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run.
A build is successful when all validation checks pass.
