Build the capitalization threshold, standing audit trail, computed depreciation, required disposal
date, and inventory-scan resolution into my school district's capital asset tracking database, using
the outcome-first method.

Outcome: an asset can never be saved on file at or under the capitalization threshold, and editing an
asset's room, custodian, department, funding source, or status always writes a row to its movement
history — both no matter how the change was made, and neither one re-blocks an asset that was already
correctly on file when the threshold last changed. An asset's book value is always current and never
goes stale. An asset can't be marked disposed without a disposal date. A scanned barcode resolves to
matched, unexpected, or duplicate on its own, and an asset nobody finds during a count is identifiable
without anything being written to its own record.

Resources — these files are inside the Open Template Scaffolds library folder; the paths below are
relative to that folder's root: templates/asset-tracking/asset-tracking-outcome-first.md (the
template), templates/asset-tracking/asset-tracking-schema.md (paired tables), standards/ (my
conventions), README.md and templates/_template-schema.md for the rest.

A build is successful when all validation checks run.
A build is successful when all validation checks pass.
