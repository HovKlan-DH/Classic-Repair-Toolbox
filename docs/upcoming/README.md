# Upcoming documentation

Pages in this folder document functionality that exists in a branch or is planned, but is
**not in a released build**. Nothing here describes what the current release can do.

Released documentation lives one level up, in [`docs/`](../README.md).

## Pages

| Page | Target version |
| --- | --- |
| _(none currently)_ | |

## When to use this folder

Use `docs/upcoming/` for **whole new topics only** — a feature that will get its own page.

If the change is an addition to an existing, already-released page, do **not** split it out.
Mark the affected section inline with the same callout instead, so related information stays
together:

```markdown
> [!NOTE]
> **Unreleased.** Planned for v2.6.
> This section describes functionality that is not available in the current release.
```

## Page format

Every page in this folder starts with an H1 followed by an "Unreleased" callout. The target
version is **required**; a tracking issue link is optional.

```markdown
# Plugin support

> [!NOTE]
> **Unreleased.** Planned for v2.6.
> This page describes functionality that is not available in the current release.
```

Nothing in this folder may be linked from [`docs/README.md`](../README.md) as though it were
released. It is linked only from the clearly labelled "Upcoming" section at the bottom of that
page, or not at all.

## Promotion process — when a feature ships

1. Move the page up a level:

   ```
   git mv docs/upcoming/<feature>.md docs/<feature>.md
   ```

2. Move any images used **only** by that page from `docs/upcoming/images/` to `docs/images/`,
   and update the references in the page.
3. Remove the "Unreleased" callout from the page.
4. Add the page to the table of contents in [`docs/README.md`](../README.md), in the section
   where it belongs.
5. Remove its entry from the **Pages** table above.
6. Check for links elsewhere in `docs/` that still point at the old `./upcoming/<feature>.md`
   path and repoint them:

   ```
   grep -rn "upcoming/<feature>" docs/
   ```
