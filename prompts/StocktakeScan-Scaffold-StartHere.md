Build scan-driven stocktake counting into my Access database using the scaffold route.

Outcome: every scan is recorded and nothing scanned is silently dropped, duplicate scans of the same
item are recognized and not double-counted, a barcode that matches nothing known is held for review
rather than guessed at or discarded, and any product whose count strays outside its tolerance —
including one nobody scanned at all — is flagged for review.

Scaffold route: the template ships the procedures already written, with the house-specific parts
marked for you to fill in under my standards. (Outcome-first is the other route: the template sets
the finished condition and no code, and you choose the structure yourself.)

Resources, relative to the Open Template Scaffolds folder:
templates/stocktakescan/stocktake-scan-scaffold.md (the template),
templates/stocktakescan/stocktake-schema.md (paired tables), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run and pass.
