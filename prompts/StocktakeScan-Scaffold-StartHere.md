Build scan-driven stocktake counting into my Access database using the scaffold approach.

Outcome: every scan is recorded and nothing scanned is silently dropped, duplicate scans of the same
item are recognized and not double-counted, a barcode that matches nothing known is held for review
rather than guessed at or discarded, and any product whose count strays outside its tolerance —
including one nobody scanned at all — is flagged for review.

Resources — these files are inside the Open Template Scaffolds library folder; the paths below are
relative to that folder's root: templates/stocktakescan/stocktake-scan-scaffold.md (the template),
templates/stocktakescan/stocktake-schema.md (paired tables), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run.
A build is successful when all validation checks pass.
