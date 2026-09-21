Build scan-driven stocktake counting into my Access database using the outcome-first route.

Outcome: every scan is recorded and nothing scanned is silently dropped, duplicate scans of the same
item are recognized and not double-counted, a barcode that matches nothing known is held for review
rather than guessed at or discarded, and any product whose count strays outside its tolerance —
including one nobody scanned at all — is flagged for review.

Outcome-first route: the template sets the finished condition and ships no code; you choose the
structure and write it, under my standards. (Scaffold is the other route: the procedures come
written, with the house-specific parts marked for you to fill in.)

Resources, relative to the Open Template Scaffolds folder:
templates/stocktakescan/stocktake-scan-outcome-first.md (the template),
templates/stocktakescan/stocktake-schema.md (paired tables), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run and pass.
