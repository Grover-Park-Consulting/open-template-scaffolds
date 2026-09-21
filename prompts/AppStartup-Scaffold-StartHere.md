Build application startup and back-end relinking into my Access database using the scaffold route.

Outcome: my front end opens by itself and finds its data file — reconnecting to it without saying
anything if it moved, or asking me where it went if it can't find it on its own, and refusing a file
that isn't this application's data. Where the data file is kept behind a database password, I'm asked
for it once, while this is being set up, and never again after that. Once the data can be reached, my
application's shared and per-person folders are confirmed or created, and then my startup screen opens.

Scaffold route: the template ships the procedures already written, with the house-specific parts
marked for you to fill in under my standards. (Outcome-first is the other route: the template sets
the finished condition and no code, and you choose the structure yourself.)

Resources, relative to the Open Template Scaffolds folder:
templates/app-startup/app-startup-scaffold.md (the template), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run and pass.
