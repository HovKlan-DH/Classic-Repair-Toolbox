# Components in scope

← [Back to Workbooks](Workbooks) · Part of [the worklog editor](Workbooks-The-worklog-editor)

Two checklists in the worklog editor tie a worklog to the actual components on the board.

---

## Components in scope

**Which components this worklog is about.** `U19`, `C64`, `R12`.

When you draw a worklog's area, the app works out which components the rectangle **touches** — using
the same component highlight rectangles the Schematics tab draws — and offers exactly those. For a
new worklog they all start **ticked**, on the assumption that you drew the rectangle around them
because they are what you meant.

Untick anything that only happened to fall inside the rectangle. A generous rectangle around a chip
will usually catch a couple of neighbouring passives.

**All** and **None** buttons sit in the section header for the common cases.

Being in scope is what makes a worklog findable by component later — the scope is searched by the
["Find a previous repair" box](Workbooks-Searching), and it is what ranks a worklog to the top when you
[attach an oscilloscope capture](Workbooks-Attaching-oscilloscope-captures) taken on that component.

---

## Components completed

**Which of the in-scope components you have actually dealt with.**

This list offers exactly the components currently ticked into scope, and nothing else. Untick a
component from the scope and it disappears from here too — a component the worklog no longer covers
cannot be one it has completed.

It exists for jobs that are a list of the same task repeated: "replace every electrolytic",
"reflow all the socket pins", "check every rail". One worklog, a dozen components, ticked off as you
go.

Everything starts unticked. A newly added component is work still to do.

**All** and **None** are here too.

---

## When the section is not shown

The checklists only appear when the app can work out what your area covers. That needs board data
loaded *and* the component highlight rectangles for **this worklog's own schematic**.

There are three distinct outcomes, and the difference matters:

| Situation | What you see | What happens to your saved scope |
| --- | --- | --- |
| Scope known, area covers components | The checklist | Saved normally |
| Scope known, area covers nothing | *"No components in this area"* | Saved as empty |
| Scope **unknown** | The section is hidden entirely | **Left untouched** |

That last row is the important one. If the app cannot tell what the area covers — the board data has
not finished loading, or a region filter has hidden the highlights for this schematic — it hides the
section rather than showing an empty one, and saving the worklog **does not wipe** the components
you had already put in scope.

If you open a worklog and expect the checklist but do not see it, let the board finish loading and
reopen it, or check that the region filter is not hiding the components.

---

## Regions

On boards that have several **regions** (board revisions or variants), the highlight rectangles the
scope is built from follow the region filter. A component hidden by the current region is not
offered.

Your saved scope is never altered by this — a worklog scoped to a component that the current region
hides keeps that component; it simply is not offered for editing while the region hides it.

---

**Next:** [Marking areas on a schematic](Workbooks-Marking-areas-on-a-schematic) · [Searching](Workbooks-Searching)
