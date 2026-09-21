Build the active-official check, team and game-time validation, derived play level and age, the
applicable pay rate, and the photo file/folder handling into my sports officiating database using the
scaffold approach.

Outcome: an assignment can never reference an official who isn't active, a game can never be saved
against itself or with its end time before its start, a game's play level always reflects its home
team with nothing to recompute by hand, an official's age is always current and never stored, the pay
rate resolved for an assignment is always the one in force on the game's date, and picking a photo for
an official always lands in one shared, confirmed folder under a name that can't collide with anyone
else's.

Resources — these files are inside the Open Template Scaffolds library folder; the paths below are
relative to that folder's root: templates/scheduling-assignment/officiating-assignment-scaffold.md
(the template), templates/scheduling-assignment/officiating-assignment-schema.md (paired tables),
standards/ (my conventions), README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run.
A build is successful when all validation checks pass.
