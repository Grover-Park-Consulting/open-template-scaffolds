Build the game assignment form for my sports officiating database using the form-spec route.

Outcome: a form where I schedule a game and staff its crew — the crew is rows I add, so a third
official is one more row and never a change to the form. Assigning someone already on that game, or
someone who isn't active, is refused with a plain message rather than a database error. A game with
the same team on both sides, or with mismatched play levels, is caught when I save. Adding, saving,
deleting and moving between records all work.

Form-spec route: the template sets out the form's controls and everything the form has to do, and
stops at function rather than looks — a working, unstyled form is a pass, and the styling is mine.
This is the only route for a form; the scaffold and outcome-first routes build the logic behind
one, not the form itself.

Build the paired scaffold first if it isn't already in place — the form's checks and its pay-rate
and photo-folder lookups call procedures that live there.

Resources, relative to the Open Template Scaffolds folder:
templates/scheduling-assignment/officiating-assignment-form.md (the template),
templates/scheduling-assignment/officiating-assignment-schema.md (paired tables),
templates/scheduling-assignment/officiating-assignment-scaffold.md (the paired scaffold),
standards/ (my conventions), README.md and templates/_template-schema.md for the rest.

A build is successful when the form opens, every control the template lists is bound as it says,
and every one of the form's features works when I try it.
