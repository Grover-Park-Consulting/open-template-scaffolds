Build error logging into my Access database using the scaffold route.

Outcome: every unhandled VBA error produces exactly one record, and the log can never refuse to
accept it — the error's number, description, and failing line are captured before anything else in
the logger can overwrite them, a description too long to fit is shortened rather than dropped, and
the person at the keyboard is never told something was saved when it wasn't.

Scaffold route: the template ships the procedures already written, with the house-specific parts
marked for you to fill in under my standards. (Outcome-first is the other route: the template sets
the finished condition and no code, and you choose the structure yourself.)

Resources, relative to the Open Template Scaffolds folder:
templates/errors/error-logging-scaffold.md (the template),
templates/errors/error-logging-schema.md (paired tables), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run and pass.
