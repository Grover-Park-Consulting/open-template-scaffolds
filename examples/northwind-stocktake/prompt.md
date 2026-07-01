# Worked Example — The Prompt (Northwind Stocktake)

This is the prompt from [`prompts/build-table-schema.md`](../../prompts/build-table-schema.md),
filled in for a real request: adding scan-driven stocktake to a Northwind-based inventory database.
See [`output.md`](output.md) for exactly what the AI produced from it.

---

```text
- Build: add stocktake scanning to our Northwind inventory database
- Standards: use my own standards — match the existing Northwind database (no tbl/tlkp
  prefixes; audit columns AddedBy / AddedOn / ModifiedBy / ModifiedOn maintained by data
  macros, like the rest of Northwind)
- Who this is for: Northwind Traders' warehouse team. They take stock a few times a year with
  handheld barcode scanners — one person walks the shelves scanning each product, and the scans
  add up to the counted quantity automatically. They want each count compared against what the
  system expected on hand, the shrinkage shown as a percentage, and anything over 5% flagged for
  someone to review. A few high-value products have a tighter tolerance that should override the
  5% default.
- Extra options: none
```

*(The "Instructions to the AI" section from the prompt file is pasted along with the box above.
It is reproduced below, unchanged, so this example is self-contained.)*

---

### Instructions to the AI

You are generating a set of tables for the person and project described above. Work strictly from
the matched template and the standards layer — do not invent content beyond what the template
defines and what I've supplied under "Who this is for."

1. **Match and read** the template that fits the "Build" line, and read the **standards** (the
   library default unless I named my own).
2. **Honor the template** — its tables, fields, relationships, and business rules are decisions
   already made. Carry them through unless my notes override a specific point.
3. **Apply the standards** — naming, audit columns, and the error-handling pattern come from the
   standards layer, never from the template itself. Apply the field-naming rules (no bare reserved
   or ambiguous words — `Status` becomes `<Entity>StatusID`).
4. **Surface every declared house assumption** the template carries. List each one and ask me to
   confirm or change it before finalizing — these are choices I should see, not inherit silently.
5. **Fold in my specifics and any extra options** I asked for. If I named an option, pull it from
   the template's options list into the live result.
6. **Ask, don't guess.** If the template and my notes leave something genuinely undecided, ask me —
   don't fill the gap with invented content.

Then produce the result in two parts:

1. **A diagram** — a fenced `mermaid` `erDiagram` showing the tables, their keys, how they relate
   (one-to-many, etc.), and how they connect to any existing tables the template builds on.
2. **A field list** — one `| Field | Type | Key / Req | Purpose & rules |` table per table, plus
   each table's indexes, any derived (not stored) values, and the audit columns from the standards
   layer.

Present it for my review. I'll approve it or tell you what to change — this is never finished until
I say so.
