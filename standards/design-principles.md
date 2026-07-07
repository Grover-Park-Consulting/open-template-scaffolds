# Design Principles — GPC Default Standards Layer

> **GPC default; fork-and-replace.** The reasoning behind the specific rules in this standards
> layer — the principles that decide how generated VBA is shaped. A `vba-scaffold` (or any code
> generated alongside a template) defers to these via `standards_layer: [design-principles]`. The
> concrete rules in the sibling files (`query-style.md`, `error-handling.md`, `form-conventions.md`,
> `naming-conventions.md`) are these principles made specific. A forked practice may swap or extend
> them. Applies to procedural Microsoft Access / VBA.

---

VBA is a procedural language. It has class modules, but not full object-oriented features: no
inheritance, no clean interfaces. Generated code is procedural by default.

The principles below come from object-oriented programming, but they are not really about objects.
They are about keeping code simple to read and safe to change, and they apply just as well to an
ordinary procedure. They are the reasoning behind the specific rules in the sibling standards files.
When a rule seems to work against one of these principles, trust the principle. The rule is being
misapplied.

## Single Responsibility — one job per procedure

A procedure should do one job. A module should hold procedures that share one purpose.

The test is simple. Describe what a procedure does in one sentence. If the honest description needs
an "and" — it reads the input, and checks it, and builds the SQL, and refreshes the form — then it is
doing several jobs, and each of them belongs in its own procedure.

This goes wrong most often in event handlers. An event handler should decide what happens and call
the procedure that does it. It should not carry the work itself. Keep the code that decides apart
from the code that acts.

The payoff is practical. A short procedure that does one thing is easy to read, to test on its own,
and to reuse. A long procedure that does five things hides its mistakes among them, and no one part
can be reused without the other four. When every procedure does one job, there is no overloaded
procedure left to bolt an exception onto.

## Separation of Concerns — keep different kinds of work apart

Single Responsibility says a procedure does one job. Separation of Concerns is the same idea one
level up: different *kinds* of work belong in different places.

A record-finding form holds at least three kinds of work. Building a SQL string is one. Deciding
which record to show is another. Reading and writing the controls on the form is a third. Each kind
should sit in its own procedure, and often its own module. When they are tangled in one place, a
change to the SQL can break the screen handling, because they share the same code.

The sign of mixed concerns: you open a procedure to change how something looks and find yourself
editing a SQL string; or you open it to fix a query and find form controls. Those are two concerns
that should have been apart.

Keeping them apart also lets you change one without touching the others. How SQL is written is a
house-style decision (`query-style.md`); which record to show is application logic. Separate, a
change to one is a change in one place.

## Encapsulation — hide the details behind a clear surface

A procedure should be usable from its name and its parameters alone. A caller should not need to know
how it works inside.

State is where this matters most. When several procedures share a value, that value should have one
owner. Pass it as a parameter, or hold it in one place with a typed accessor. Do not scatter it as a
global that many procedures each read and interpret on their own.

The sign of a problem: the same global value read in three different procedures, each deciding for
itself what it means. The value has no owner, and the day its meaning changes you must find every
reader and hope you missed none. (A record finder that reads its A–Z letter from a global TempVar in
three places, while passing its category as a parameter, is moving the same kind of value two
different ways for no reason.)

Two rules in this layer already serve this: prefer TempVars with typed wrappers over Public module
variables for shared state (`query-style.md`), and keep helper procedures Private so a caller sees
only the few entry points it needs.

## High Cohesion and Low Coupling — things that belong together, few connections between them

Two related ideas. Cohesion is about what goes inside a module: its procedures should belong
together, serving one purpose. Coupling is about the connections between modules: a module should
depend on as few others as possible, and only through their stated surface.

Low cohesion looks like a module named for a grab-bag, or one task split needlessly across several
modules so you jump around to follow it. High cohesion is a module you can describe in a few words,
holding the procedures that carry out that one thing.

High coupling looks like a change inside one procedure forcing changes in several others that reached
into its internals. Low coupling is a procedure other code calls through a stable signature: you can
rewrite its inside freely as long as the signature holds. (A shared query-rewriting helper called by
three different features: because they all call it the same way, its inside can be made safe without
touching any of them.)

The test for coupling: if you change how a procedure works inside, how many other procedures must
change too? The answer should be none.

## DRY — one home for each piece of knowledge

DRY stands for Don't Repeat Yourself. Any single fact about how the application works should live in
exactly one place in the code.

The usual sign is the same block of logic copied into two procedures. The danger is not the typing;
it is the day you fix one copy and forget the other. Now the application behaves two ways, and the
difference stays invisible until it bites.

When you find the same logic in two places, move it to one procedure and call that from both. (A
finder that builds the same "show everything" pick-list SQL in two places, with small differences
between the copies, has two paths that can drift apart; moved to one builder, they cannot.)

DRY has a limit worth stating: do not force two things together just because they look alike today.
Merge them when they are the same fact expressed twice, not when they are two facts that happen to
resemble each other for now.

## Strong Contracts — say exactly what a procedure takes and returns

A procedure's signature is a promise. State it fully. Declare the type of every parameter, mark each
ByVal or ByRef on purpose, and give every Function an explicit return type. An untyped parameter is a
Variant, and a Function with no As clause returns a Variant — so the compiler can no longer catch a
wrong type, and the reader can no longer tell what was intended.

The sign of a weak contract: a Function you cannot understand without reading its body, because its
name and signature do not tell you what it returns. (A finder builder declared with no return type
still returns a SQL string, but nothing in its signature says so. Adding As String makes the promise
visible and lets the compiler enforce it.)

This also decides Sub versus Function. A Function returns a value and is called for that value. A Sub
performs an action and returns nothing. A Function that returns nothing, or one whose result is
always ignored, is usually a Sub in the wrong clothes. (A shared helper that is a Function returning
the same name it was handed did its work and returned nothing useful — it was really a Sub.)

---

*Worked example:* `templates/library/record-finder-scaffold.md` is a `vba-scaffold` shaped by these
principles throughout — one job per procedure, the query-rewriting helper called through a stable
signature, the "show everything" row built in one place, and every builder typed.
