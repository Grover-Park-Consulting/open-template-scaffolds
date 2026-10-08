# Start-up form text

**Who reads this:** whoever maintains the wording of the start-up form. The form (`start-form.ps1`,
in this folder) reads this file each time it runs and shows it to the developer, so editing a line
here changes the window with no code change.

**How it is read:** each `##` heading below names one piece of the window; the text under it is that
piece. Keep the heading names exactly. Cards are numbered `Card 1`, `Card 2` and so on, each with a
`heading` and a `text`; add or remove whole cards as needed, keeping the numbers in order. Anything
above the first `##` heading is ignored. Bold marks (`**`) are removed when the window is drawn. In
the countdown line, `{time}` is replaced by the minutes and seconds left.

The window is shown once at the start of a run. When the developer presses the button, or the
countdown runs out, it moves to the taskbar, and the developer can reopen it there for the rest of
the run. The pieces marked "when reopened" are used then instead.

## Banner, line 1

O P E N   T E M P L A T E   S C A F F O L D S

## Banner, line 2

design it, check it, then build it

## Window title (the large line above the cards)

Before you answer the first question

## Window title, when reopened

What you were told at the start

## Card 1: heading

You decide where this run ends: at an approved design, or at a finished build

## Card 1: text

Runs can end with an approved design, or they can end with something built and checked. Building depends on your assistant being able, and permitted, to open and run things in your Access database. Your assistant tells you whether it can do so in its first message. If it can, you decide whether to allow it to do so.

## Card 2: heading

What this template is

## Card 2: text

A template holds design decisions that we already worked out and tested for you. Your assistant reads one together with your own standards, then shows you a design to approve or change. The design decision is yours.

## Card 3: heading

Nothing is created until you say so.

## Card 3: text

You see the design first: a table diagram and the detail of each field. Nothing is created or changed in your database until you approve it and tell your assistant to go ahead.

## Card 4: heading

Your open work is left alone

## Card 4: text

If something is built, it is built and checked in a separate copy of your database, not in a file you have open. A window you already have open is never touched.

## Card 5: heading

How you will see the work

## Card 5: text

Everything you need to decide arrives as a question. Everything else goes into a **build record**, saved next to the result, that records what was checked and what surprised your assistant. How much of the assistant's work scrolls past on screen depends on the assistant you use. That changes nothing about what you get.

## Line above the button

When you are ready, press the button and go back to your assistant. This window moves to your taskbar, where you can open it again at any time during the run. You can also ask your assistant to show it again.

## Countdown line

Take your time. If you have not pressed the button in {time}, this window moves to your taskbar by itself.

## Button label

Continue to the Build

## Button label, when reopened

Set aside
