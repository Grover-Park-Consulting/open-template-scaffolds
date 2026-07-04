# README fixture

Any `README.md` must be skipped by `iter_templates` — this file has no front-matter
on purpose; if it is ever yielded, tests fail loudly.
