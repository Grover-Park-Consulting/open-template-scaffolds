Build the sort-title derivation and volume/set validation into my library catalog database using the
outcome-first route.

Outcome: a publication's sort title always matches its title with leading noise words removed, set
correctly the moment a title is added and kept correct whenever it's edited afterward — the route I
choose for this decides whether that coverage extends to every way the table can be edited or only to
my entry form, and I'm told which before anything is built. A publication's volume and set fields can
never be saved in an inconsistent combination, whatever the route the save was made through.

Outcome-first route: the template sets the finished condition and ships no code; you choose the
structure and write it, under my standards. (Scaffold is the other route: the procedures come
written, with the house-specific parts marked for you to fill in.)

Resources, relative to the Open Template Scaffolds folder:
templates/library/catalog-outcome-first.md (the template),
templates/library/catalog-schema.md (paired tables), standards/ (my conventions),
README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run and pass.
