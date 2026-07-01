---
template: _materialization
title: GPC Template Library — Materialization (table-schema + form-spec)
domain: _meta
type: spec
version: 0.1.0
status: draft
---

# Materialization — building real artifacts from a template

This is a **format/process reference** (like `_template-schema.md`): it defines how a template's
approved design becomes a real, buildable artifact, and **proves each mapping by hand** so the MCP
generator (phase B3) is never built against an unproven mapping. It is meta, not a template
(`type: spec`, `domain: _meta`). Two builds are covered:

- **Table-schema → tables** — Access local tables via a VBA DAO `Sub`, or SQL Server via `CREATE TABLE` DDL.
- **Form-spec → a form** — importable Access form text (`SaveAsText` / `LoadFromText`).

---

## Table-schema build — Access (VBA-DAO) or SQL Server

After a `table-schema` design is approved and the developer asks you to build it, **ask which platform**
and generate the matching artifact. Both carry the tables, fields (with their **comments**), keys,
indexes, relationships, and lookup **seed rows**.

### SQL Server → DDL

`CREATE TABLE` statements with primary keys, foreign keys, indexes, the lookup tables, and `INSERT` seed
rows. Field comments become inline comments or extended properties.

### Access (ACE) → a VBA `Sub` using DAO

Access local tables are **not** built from a `CREATE TABLE` query — they're built in a VBA `Sub` with
DAO, because that is the only way to carry a field's **`Description`** (the integral comments) and to set
AutoNumber, indexes, and relationships fully. This pattern is **proven by running it** (see the four
rules below); generalize it to any `table-schema`:

```vba
Option Compare Database
Option Explicit

' Build the <domain> tables in the current database, using DAO.
'
' RUN THIS FROM A TRUSTED LOCATION. Outside one, Access silently disables VBA
' and nothing is created.
'
' The errHandler here is self-contained (a message box) so the Sub runs with no
' outside dependencies. A shop with a central error logger swaps the errHandler
' block for its own (per standards/error-handling.md).

Public Sub BuildCatalogTables()
    Dim db  As DAO.Database
    Dim tdf As DAO.TableDef
    Dim rel As DAO.Relation
    Dim fld As DAO.Field

    On Error GoTo errHandler
    Set db = CurrentDb

    ' Safety guard: never overwrite existing tables.
    If TableExists(db, "tlkpMediaType") Or TableExists(db, "tblPublication") Then
        MsgBox "One or more target tables already exist. Build aborted.", vbExclamation
        GoTo Cleanup
    End If

    ' ---- tlkpMediaType (lookup) ----
    Set tdf = db.CreateTableDef("tlkpMediaType")
    AddField tdf, "MediaTypeID", dbLong, 0, True, True        ' AutoNumber primary key
    AddField tdf, "MediaTypeName", dbText, 255, True, False
    AddField tdf, "SortOrder", dbLong, 0, False, False
    AddPrimaryKey tdf, "MediaTypeID"
    db.TableDefs.Append tdf
    ' Descriptions must be set AFTER the table is appended (else error 3219).
    SetDesc db, "tlkpMediaType", "MediaTypeID", "Surrogate key"
    SetDesc db, "tlkpMediaType", "MediaTypeName", "Media-type label"
    SetDesc db, "tlkpMediaType", "SortOrder", "Display order"

    ' ---- tblPublication (entity) ----
    Set tdf = db.CreateTableDef("tblPublication")
    AddField tdf, "PublicationID", dbLong, 0, True, True      ' AutoNumber primary key
    AddField tdf, "PublicationTitle", dbMemo, 0, True, False
    AddField tdf, "MediaTypeID", dbLong, 0, False, False      ' foreign key to tlkpMediaType
    AddField tdf, "MultiVolumeSet", dbBoolean, 0, True, False
    AddField tdf, "AddedOn", dbDate, 0, True, False, "Now()"  ' audit default = field property, not DDL
    AddPrimaryKey tdf, "PublicationID"
    AddIndex tdf, "MediaTypeID", False
    db.TableDefs.Append tdf
    SetDesc db, "tblPublication", "PublicationID", "Surrogate key"
    SetDesc db, "tblPublication", "PublicationTitle", "The full title, untruncated"
    SetDesc db, "tblPublication", "MediaTypeID", "Format / media type (FK to tlkpMediaType)"
    SetDesc db, "tblPublication", "MultiVolumeSet", "True = a multi-volume set"
    SetDesc db, "tblPublication", "AddedOn", "When the row was created (defaults to Now())"

    ' ---- relationship: tlkpMediaType (1) -> (many) tblPublication ----
    Set rel = db.CreateRelation("tlkpMediaTypetblPublication", "tlkpMediaType", "tblPublication")
    Set fld = rel.CreateField("MediaTypeID")
    fld.ForeignName = "MediaTypeID"
    rel.Fields.Append fld
    db.Relations.Append rel

    ' ---- seed rows for the lookup ----
    db.Execute "INSERT INTO tlkpMediaType (MediaTypeName, SortOrder) VALUES ('Book', 1)", dbFailOnError
    db.Execute "INSERT INTO tlkpMediaType (MediaTypeName, SortOrder) VALUES ('Periodical', 2)", dbFailOnError
    db.Execute "INSERT INTO tlkpMediaType (MediaTypeName, SortOrder) VALUES ('Map', 3)", dbFailOnError

    Application.RefreshDatabaseWindow
    MsgBox "Catalog tables built successfully.", vbInformation

Cleanup:
    On Error Resume Next
    Set fld = Nothing: Set rel = Nothing: Set tdf = Nothing: Set db = Nothing
    Exit Sub

errHandler:
    MsgBox "Build failed: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Sub AddField(tdf As DAO.TableDef, ByVal sName As String, ByVal lngType As Long, _
                     ByVal lngSize As Long, ByVal booRequired As Boolean, ByVal booAutoNum As Boolean, _
                     Optional ByVal sDefault As String = "")
    Dim fld As DAO.Field
    If lngSize > 0 Then
        Set fld = tdf.CreateField(sName, lngType, lngSize)
    Else
        Set fld = tdf.CreateField(sName, lngType)
    End If
    If booAutoNum Then
        fld.Attributes = dbAutoIncrField        ' AutoNumber = Long + auto-increment
    ElseIf booRequired Then
        fld.Required = True
    End If
    ' A field default is a property set BEFORE append (e.g. sDefault = "Now()") --
    ' never a DDL DEFAULT clause; the DAO/ANSI-89 engine rejects DEFAULT in CREATE TABLE.
    If Len(sDefault) > 0 Then fld.DefaultValue = sDefault
    tdf.Fields.Append fld
End Sub

Private Sub SetDesc(db As DAO.Database, ByVal sTable As String, ByVal sField As String, ByVal sDesc As String)
    Dim prp As DAO.Property
    Set prp = db.TableDefs(sTable).Fields(sField).CreateProperty("Description", dbText, sDesc)
    db.TableDefs(sTable).Fields(sField).Properties.Append prp
End Sub

Private Sub AddPrimaryKey(tdf As DAO.TableDef, ByVal sField As String)
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    Set fld = idx.CreateField(sField)
    idx.Fields.Append fld
    tdf.Indexes.Append idx
End Sub

Private Sub AddIndex(tdf As DAO.TableDef, ByVal sField As String, ByVal booUnique As Boolean)
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Set idx = tdf.CreateIndex("idx" & sField)
    idx.Unique = booUnique
    Set fld = idx.CreateField(sField)
    idx.Fields.Append fld
    tdf.Indexes.Append idx
End Sub

Private Function TableExists(db As DAO.Database, ByVal sName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = db.TableDefs(sName)
    TableExists = (Not tdf Is Nothing)
End Function
```

**Four rules that make the generated Sub actually run** — each learned by running it against a real
database:

1. **Trusted Location.** The Sub must run from an Access **Trusted Location**; outside one, Access
   silently disables VBA and nothing is created (a description-less failure). Always tell the developer.
2. **Descriptions after append.** Setting a field's `Description` during field-build throws **error
   3219**; create the `Description` property in a **second pass, after `db.TableDefs.Append`**.
3. **AutoNumber** is a `dbLong` field with `Attributes = dbAutoIncrField` (and no explicit `Required`).
4. **Defaults are a field property, never DDL.** Set `fld.DefaultValue` (e.g. `"Now()"`) **before**
   appending the field. Do **not** emit a `CREATE TABLE … DEFAULT …` statement — the DAO/ANSI-89
   engine rejects the `DEFAULT` keyword outright (proven: even `DEFAULT 'literal'` fails with a syntax
   error). This is the exact trap that turns a generated ACE-DDL build into "Syntax error in CREATE
   TABLE statement."

Naming, audit columns, and types follow the active standards; the `errHandler` is the standards-layer
one (the **dependency-free message-box default** unless `error-handling.md` specifies a central logger).

---

*The rest of this document covers **form-spec** materialization.*

---

## Layout fidelity — a limitation of the approach

The generated layout is a **functional default**, not a reproduction of a real 2D form design —
multiple columns, multiple controls per line, landscape balance. A structural spec deliberately avoids
specifying detailed layouts, partly because of the variability of controls, design patterns and user
preferences, and partly in recognition of the fact that the adopter will implement their own preferred
layout in any case. **This is a limit of capturing forms as structural specifications, not a limitation
of the AI** — implementation-specific information can't be contained in the spec to begin with. When it
materializes a form, the AI **says so up front**: *"This is a functional default layout; arranging it
into your columns and lines is your styling pass."*

---

## Mapping rules

1. **Form shell** — `record_source` → `RecordSource`; `title` → `Caption`; default view.
2. **Regions → sections** — Form Header / Detail / Form Footer → the Access form's sections.
3. **Control row → control block** — name + type → `Begin <Type>`; `Bound to` → `ControlSource`; a
   lookup combo → `RowSource` (+ hidden bound column); Boolean → `CheckBox`; a subform → `SourceObject`
   + `LinkMasterFields` / `LinkChildFields`.
4. **Default-layout rule** — single-field controls (textbox, combo, checkbox) stack vertically in
   inventory order (label-left / control-right, fixed row pitch). **Subform controls are never
   interleaved into that stack — they are placed to the *right* of it, or *below* it when width doesn't
   allow.** That is the *one* layout judgment the default makes; beyond it, no 2D optimization is
   attempted (see Layout fidelity).
5. **Code-behind** — event handlers wired to the named framework helpers, each with the standard
   `errHandler` block (`error-handling.md`; line numbering per the house policy). Helpers are **called,
   not defined**.

---

## Hand-validation — publication form (focused, importable fragment)

A representative slice of `templates/library/publication-form.md` materialized to Access text: the form
shell, a Detail section with a label + textbox, a lookup combo, and a subform placed to the right of the
single-field stack (rule 4), plus one framework-wired event handler.

```text
Begin Form
    RecordSource ="qryPublication_frm"
    Caption ="Library Catalog — Publication"
    Begin
        Begin Label
            Name ="lblPublicationTitle"
            Caption ="Publication Title"
            Left =120  Top =120  Width =2160  Height =480
        End
        Begin TextBox
            Name ="txtPublicationTitle"
            ControlSource ="PublicationTitle"
            Left =2400 Top =120 Width =6000 Height =960     ' Memo -> taller, multi-line
        End
        Begin ComboBox
            Name ="cboMediaTypeID"
            ControlSource ="MediaTypeID"
            RowSource ="SELECT MediaTypeID, MediaTypeName FROM tlkpMediaType ORDER BY SortOrder"
            ColumnCount =2
            ColumnWidths ="0;2880"
            Left =2400 Top =1200 Width =6000 Height =480
        End
        Begin SubForm                                       ' [rule 4] subform to the RIGHT of the stack
            Name ="sfrmPublication_Creator"
            SourceObject ="sfrmPublication_Creator"
            LinkMasterFields ="PublicationID"
            LinkChildFields ="PublicationID"
            Left =10680 Top =120 Width =6000 Height =1440
        End
    End
End
```

```vba
' CodeBehindForm
Private Sub cboSelectPublication_AfterUpdate()
100   On Error GoTo errHandler
110   With Me
120       .RecordSource = RefreshSQLWhere(.RecordSource, _
              IIf(.cboSelectPublication = 0, "1=1", "PublicationID = " & .cboSelectPublication), _
              "", "PublicationSortTitle")
130   End With
Cleanup:
140   Exit Sub
errHandler:
150   Call GlblErrMsg(sFrm:=Application.VBE.ActiveCodePane.CodeModule, sCtl:="cboSelectPublication_AfterUpdate")
160   Resume Cleanup
End Sub
```

*(The generator also emits the boilerplate `SaveAsText` headers — `Version` / `Checksum` / `NameMap` /
`PrtMip`. This fragment shows the meaningful mapping — control blocks, default positions, the
subform-to-the-right rule, and one framework-wired handler — which is what the hand-validation
confirms. A byte-perfect importable file is the generator's job in B3.)*

---

## Alternative path — build live via MCP

The same mapping drives the Access Explorer MCP `create_form` / `create_control` tools: instead of
emitting text for import, the generator creates the form and its controls directly, applying the same
default-layout rule and wiring the same code-behind. The `form-spec` markdown remains the source of
truth; both paths are generated targets.
