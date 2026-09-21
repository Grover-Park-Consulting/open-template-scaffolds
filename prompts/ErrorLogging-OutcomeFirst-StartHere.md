Build error logging into my Access database using the outcome-first route.

Outcome: every unhandled VBA error produces exactly one record, and the log can never refuse to
accept it — the error's number, description, and failing line are captured before anything else in
the logger can overwrite them, a description too long to fit is shortened rather than dropped, and
the person at the keyboard is never told something was saved when it wasn't.

Outcome-first route: the template sets the finished condition and ships no code; you choose the
structure and write it, under my standards. (Scaffold is the other route: the procedures come
written, with the house-specific parts marked for you to fill in.)

Resources, relative to the Open Template Scaffolds folder:
templates/errors/error-logging-outcome-first.md (the template),
templates/errors/error-logging-schema.md (paired tables), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run and pass.
