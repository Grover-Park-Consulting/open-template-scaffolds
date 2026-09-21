Build the publications entry form for my library catalog database using the form-spec route.

Outcome: a form I can actually work in — I find the one title I want by category, by first letter,
or by keyword, and the form jumps to it, with a "show all" choice in every mode. I can add a
publisher, author, or genre without leaving the form. Anything invalid is flagged when I save and
cleared when I move on. Adding, saving, deleting and moving between records all work.

Form-spec route: the template sets out the form's controls and everything the form has to do, and
stops at function rather than looks — a working, unstyled form is a pass, and the styling is mine.
This is the only route for a form; the scaffold and outcome-first routes build the logic behind
one, not the form itself.

Build the paired record-finder scaffold first if it isn't already in place — the form's pick-list
and jump-to-record call procedures that live there.

Resources, relative to the Open Template Scaffolds folder:
templates/library/publication-form.md (the template),
templates/library/catalog-schema.md (paired tables),
templates/library/record-finder-scaffold.md (the paired scaffold), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when the form opens, every control the template lists is bound as it says,
and every one of the form's features works when I try it.
