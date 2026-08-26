# Worked Example — The Prompt (Northwind Stocktake)

**Who reads this:** anyone who wants to see a filled-in prompt before writing their own.

This is the prompt from [`prompts/BuildNewTables-StartHere.md`](../../prompts/BuildNewTables-StartHere.md),
filled in for a real request: adding scan-driven stocktake to a Northwind-based inventory database.
Open [`output.md`](output.md) to see exactly what the AI produced from it.

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
the matched template and the standards layer. Do not invent content beyond what the template
defines and what I've supplied under "Who this is for."

1. **Match and read** the template that fits the "Build" line, and read the **standards** (the
   library default unless I named my own). If you cannot see the library files, list the exact
   files you need — the matched template and the standards — and ask for them before proceeding;
   never guess at their contents.
2. **Honor the template.** Its tables, fields, relationships, and business rules are decisions
   already made. Carry them through unless my notes override a specific point.
3. **Apply the standards.** Naming, audit columns, and the error-handling pattern come from the
   standards layer, never from the template itself. Apply the field-naming rules (no bare reserved
   or ambiguous words. `Status` becomes `<Entity>StatusID`).
4. **Surface every declared house assumption** the template carries. List each one and ask me to
   confirm or change it before finalizing. These are choices I should see and make, not inherit silently.
   **Surface every declared warning the same way.** Those are hard platform limits (for example,
   Access Data Macros cannot audit Long Text fields): state each one, get my answer to whatever it
   says must be checked, and branch the build accordingly.
5. **Fold in my specifics and any extra options** I asked for. If I named an option, pull it from
   the template's **Extra Options** section into the live result.
6. **Ask, don't guess.** If the template and my notes leave something genuinely undecided, ask me. Don't
   fill the gap with invented content.
7. **No matching template? Don't stop.** Tell me which templates you considered and why each falls
   short. Then offer, as the default next step: with my approval, you'll design it from scratch
   following the same standards layer, producing the same two-part result below for the same
   review. You will ask me for the go-ahead before you begin. (Mention the alternatives in passing:
   adapting the nearest template despite the mismatch, or me refining the description.) In a
   from-scratch design, list every assumption you had to invent as **proposed assumptions** for me
   to confirm or change. I must see your choices, not inherit them silently. And once I approve
   a from-scratch design, you may offer to shape it into a template for the library. It would be
   added upon the curator's approval.

Then produce the result in two parts:

1. **A diagram** — a fenced `mermaid` `erDiagram` showing the tables, their keys, how they relate
   (one-to-many, etc.), and how they connect to any existing tables the template builds on. Under
   the diagram, add one line: "This diagram shows the shape only. Exact types and
   required/optional are in the field list below."
2. **A field list** — one `| Field | Type | Key / Req | Purpose & rules |` table per table, plus
   each table's indexes, any derived (not stored) values, and the audit columns from the standards
   layer.

Present it for my review. I'll approve it or tell you what to change. This is never finished until
I say so.

Once I approve the design and ask you to build it, **first ask whether I want Access tables (a VBA
sub) or SQL Server (DDL)**, then generate that — carrying the keys, indexes, relationships, and lookup
seed rows. For the Access version, build the tables **in the VBA sub with DAO (`CreateTableDef`), not
with `CREATE TABLE` statements** (the Access engine rejects a `DEFAULT` clause in DDL), and remind me to
**run the sub from a Trusted Location**, or Access blocks the code.
