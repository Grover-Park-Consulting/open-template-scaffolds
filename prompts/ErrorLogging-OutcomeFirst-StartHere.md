Build error logging into my Access database using the outcome-first method.

Outcome: every unhandled VBA error produces exactly one record, and the log can never refuse to
accept it — the error's number, description, and failing line are captured before anything else in
the logger can overwrite them, a description too long to fit is shortened rather than dropped, and
the person at the keyboard is never told something was saved when it wasn't.

Resources — these files are inside the Open Template Scaffolds library folder; the paths below are
relative to that folder's root: templates/errors/Error-logging-outcome-first.md (the template),
templates/errors/error-logging-schema.md (paired tables), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run.
A build is successful when all validation checks pass.
