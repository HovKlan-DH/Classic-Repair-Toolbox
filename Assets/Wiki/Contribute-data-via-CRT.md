[Wiki Home](Home)

Fix a value or add a datasheet from inside the application.

---

Spotted a wrong value, a missing part number, or do you have a datasheet, photo or scope
reading that others would benefit from? You can suggest it directly from the **Contribute**
tab - no GitHub account, no spreadsheets, no technical knowledge needed.

Your suggestion is **sent to the developer for review**. Nothing on your own machine is
changed, and nothing goes live until it has been checked. Once it is approved by application developer (maintainer of the online source data), it reaches
everyone automatically the next time the application updates its data.

## 1. Pick your board

At the top of the main window, select the **hardware** and the **board revision** you want to
improve. The "Contribute" tab always shows the board that is currently selected.

## 2. Open the "Contribute" tab

You will see all the components of the board, grouped in columns by category (ICs, capacitors,
resistors and so on). Hover a label to see its full name.

The line **"Board Excel data last revisioned"** tells you how fresh the current data is.

## 3. Click the component you want to change

Click a board label (for example `U1` or `C15`) and the **Component contribution editor**
opens in full screen. Everything in it is already filled in with the data the app has today —
you are simply correcting or extending it.

## 4. Make your changes

The editor is divided into collapsible sections. Click a heading to open it. The number in
brackets tells you how many entries it already holds.

| Section | What belongs here |
| --- | --- |
| **Component** | The basics: friendly name, technical name or value, part number, category, short description |
| **Component images** | Photos and oscilloscope readings for this component, with the scope settings used |
| **Component local files** | Files about this component — datasheets, instructions, manuals |
| **Component links** | Web links about this component |
| **Board local files** | Files covering the whole board — service and troubleshooting manuals |
| **Board links** | Web links about the whole board — repair logs, YouTube videos |

Then:

- **To change something** — just type in the box. Correct a wrong value, fill in a blank one.
- **To add something** — click the **Add new …** button at the top of the section. A blank row
  appears; fill it in.
- **To delete something** — click the **Remove** button on that row.

### Attaching a file or an image

In any row that has a **File** box, click the box and a normal file browser opens. Pick the
file from anywhere on your computer — a photo, a datasheet, a screenshot. The file name appears
in the box, images show a small preview, and the file itself is sent along with your suggestion.

The **File location** dropdown next to it says which folder the file should end up in. Leave it
as it is unless you have a reason to change it.

> **Can't find your component in the list?** Pick the closest one and simply explain in the
> comment (next step) what should be added. The developer takes it from there.

## 5. Write the mandatory comment

At the bottom, describe **what you changed, what was wrong or missing**, and anything the
developer should double-check. Please also mention **which exact board revision you have** —
boards vary, and that detail matters.

This box cannot be left empty.

## 6. Enter your email and send

Type your email address in the field at the bottom left. It is required, so the developer can
come back to you with questions, and it is remembered for next time.

Click **Send contribution update**. A progress message appears while it is sent, and you are
told when it has arrived:

> Contribution submitted successfully - thank you :-)

If something goes wrong, the message says so and you can simply try again.

Use **Cancel** to close the editor and throw your edits away. The **▲** and **▼** buttons jump
to the top and bottom of a long form.

## Worth knowing

Your own copy does not change. The suggestion goes to the developer, and once approved it reaches everyone with the next data sync — usually within a week.

## That's it

Thank you — every correction makes the data better for the next person repairing the same board.
