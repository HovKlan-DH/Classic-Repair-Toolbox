[Wiki Home](Home)

See every board exactly as everybody else sees it, without your own local changes on top.

---

When you change a board in CRT - on the "Contribute" tab (a component, a picture, the table view),
with **Mark components** or **Align KiCad** on a picture, or with **Edit this component** from a
right-click - the change goes into your own local **draft** of that board. A draft saves itself;
there is no Save button. CRT shows the board WITH your changes from then on, which is usually what
you want: you see your own work as you build it.

Sometimes you want to see the board as it really is for everybody else. That is what this setting is
for.

## Turn it on

* Go to the "Configuration" tab
* Tick **View boards as published online (hide my draft changes)**

It takes effect straight away on the board you are looking at, and on every board you open after
it. Untick it again to see your changes as usual.

## What you see while it is on

Every board shows exactly what CRT downloaded from its online source, as if you had no draft at all:

* Components, pictures and every other part you added, changed or removed are shown as they were
  before your changes.
* Pictures and other files you replaced show the downloaded version again.
* KiCad data you imported, and KiCad alignments you changed, are not used.
* The warning that the online data has changed since you started your draft is hidden.

A board that exists ONLY as your draft - a new board you created yourself - shows as an empty
board, because it is not published yet.

"Online source" means whichever source you download board data from: the stable source, or the
BETA source if you ticked **Download data from the BETA source instead of the stable source** in the
same tab.

## What it does NOT do

* **Your drafts are not touched.** Nothing is deleted, cleared or changed. Untick the setting and
  every change is back.
* **The "Contribute" tab still shows and edits your drafts** as normal - "My drafts", the component
  pane, **Table view**, **Review changes** and **Submit**. "Preview popup" there shows your draft
  whatever this setting says.
* **The "Draft" chips in the Hardware and Board drop-downs stay**, because your drafts are still
  there.

## When it is useful

* To check how a board looks for everybody else, before or after you submit a change to it.
* To compare: tick and untick it while looking at the same picture to see what your draft changes.
* To be sure a problem you see is in the downloaded data and not caused by something in your own
  draft.

## If it does not work

| What you see | Likely cause |
| --- | --- |
| Your changes have disappeared from a board | The setting is ticked. Untick it - your draft was never touched |
| A board you created yourself is empty | The setting is ticked, and a board that exists only as your draft has nothing online yet |
