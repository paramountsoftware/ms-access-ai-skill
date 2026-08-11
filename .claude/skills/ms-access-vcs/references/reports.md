# Reports Reference

> See also [Conditional Formatting](conditional-formatting.md) for format conditions on report controls.

## Critical Rules

These rules apply when editing any report `.bas` file. Violations cause silent corruption or `Error 2128` on import.

1. **TabIndex values MUST be contiguous within each section (0, 1, 2, 3...).** Gaps cause `Error 2128` on import. When removing a control, renumber all subsequent TabIndex values in that section downward to fill the gap.

2. **VBA event procedures MUST be wired up in the `.bas` file.** Adding a `Private Sub Report_Open()` (or any event) to the `.cls` code-behind is **not sufficient** — Access will silently ignore it. You must also add the corresponding event property in the `.bas` layout file. The property goes on the **object that owns the event**:
   - **Report-level events** (`Report_Open`, `Report_Load`, etc.) → property on the report object: `OnOpen ="[Event Procedure]"`, `OnLoad ="[Event Procedure]"`, etc.
   - **Control events** (`txtName_AfterUpdate`, etc.) → property on the control block: `AfterUpdate ="[Event Procedure]"`, etc.
   - **Section events** (`Detail_Format`, etc.) → property on the section block: `OnFormat ="[Event Procedure]"`, etc.

   **Without this property, the VBA code compiles but never executes.**

3. **Never hand-edit binary data blocks.** Leave `ImageData = Begin...End`, `OleData = Begin...End`, and other hex payload blocks untouched — a single byte error corrupts them. For `ConditionalFormat` and `ConditionalFormat14`, see [Conditional Formatting](conditional-formatting.md) — these must be removed entirely and replaced with VBA code.

   **Exception — `RecSrcDt`.** This is a single 8-byte little-endian IEEE-754 double holding an OLE Automation Date ("Record Source Date"): the moment Access last saved that object's record-source binding. It is a per-object cache stamp, not a checksum of `RecordSource`, and Access re-stamps it on the next design save. It is **optional** — reports with a `RecordSource` and no `RecSrcDt` import cleanly (verified against existing reports that carry a `RecordSource` with no `RecSrcDt` block). When hand-authoring a new report, either omit the block or leave a copied one in place; both are safe. Never invent a value, and never set one in the future. Copied objects legitimately share a token — Access does the same on copy-paste, so objects cloned from one original will all carry that original's stamp.

4. **Preserve Begin/End nesting exactly.** Every `Begin` must have a matching `End`. Mismatched nesting causes import failure.

5. **Control order matters.** Controls are rendered in file order (affects z-order). Don't reorder controls unless intentionally changing layering.

6. **Property values use specific formats:**
   - Numeric: `Width =1701` (units are twips: 1 inch = 1440 twips)
   - String: `Name ="Detail"` (always double-quoted)
   - Boolean flags: presence of property name implies true (e.g., `NotDefault`)
   - **String escaping uses backslash (`\"`), NOT VBA-style doubled quotes (`""`).**

7. **When removing a control block,** remove everything from `Begin ControlType` through its closing `End`, including any trailing `LayoutCached*` properties. Then renumber TabIndex values to fill the gap (Rule 1).

8. **The report `Width` must fit the printable page width.** See [Section Width vs Page Width](#section-width-vs-page-width). Re-check it after any change to control geometry or print settings.

## Report Structure

Reports use the same binary text format as forms (`Begin Report` instead of `Begin Form`). The same container structure applies: exactly ONE `Begin...End` container block holds default control styles, break level definitions, and all sections.

### Report Section Keywords

Report sections use these keywords — note that Report Header/Footer reuse the `FormHeader`/`FormFooter` keywords:

| Section | Keyword | Name Property |
|---|---|---|
| Report Header | `FormHeader` | `Name ="ReportHeader"` |
| Page Header | `PageHeader` | `Name ="PageHeaderSection"` or `Name ="PageHeader"` |
| Group Header | `BreakHeader` | `Name ="GroupHeader0"` (0-indexed) |
| Detail | `Section` | `Name ="Detail"` |
| Group Footer | `BreakFooter` | `Name ="GroupFooter0"` or `Name ="GroupFooter1"` |
| Page Footer | `PageFooter` | `Name ="PageFooterSection"` or `Name ="PageFooter"` |
| Report Footer | `FormFooter` | `Name ="ReportFooter"` |

Note: Report Header/Footer sections use the **`FormHeader`/`FormFooter` keywords** (not `ReportHeader`/`ReportFooter`). Group Header/Footer sections use **`BreakHeader`/`BreakFooter`** (not `GroupHeader0`/`GroupFooter0`).

## Sorting and Grouping

Reports define grouping via `Begin BreakLevel` blocks inside the container, placed **after** default control styles and **before** sections. Each `BreakLevel` defines one sorting/grouping level.

```
    Begin                               <- Container
        Begin Label                     <- Default styles first
            ...
        End
        Begin TextBox
            ...
        End
        Begin BreakLevel                <- Group level definitions
            GroupHeader = NotDefault     <- Show group header section
            GroupFooter = NotDefault     <- Show group footer section
            ControlSource ="OrderID"    <- Field to group by
        End
        Begin FormHeader                <- Sections follow
            Name ="ReportHeader"
            ...
        End
        Begin PageHeader
            ...
        End
        Begin BreakHeader               <- Group header section
            Name ="GroupHeader0"
            Begin
                ...controls...
            End
        End
        Begin Section
            Name ="Detail"
            ...
        End
        Begin BreakFooter               <- Group footer section
            Name ="GroupFooter0"
            Begin
                ...controls...
            End
        End
        Begin PageFooter
            ...
        End
        Begin FormFooter
            Name ="ReportFooter"
        End
    End                                 <- Closes container
```

### BreakLevel Properties

- `ControlSource ="FieldName"` — the field to sort/group by (required)
- `GroupHeader = NotDefault` — show a group header section (optional; omit for sort-only levels)
- `GroupFooter = NotDefault` — show a group footer section (optional; omit for sort-only levels)

### Multiple Break Levels

Define multiple `Begin BreakLevel...End` blocks for multi-level sorting/grouping. Only levels with `GroupHeader`/`GroupFooter = NotDefault` produce visible sections. When multiple break levels exist, add `BreakLevel =N` (0-indexed) on the `BreakHeader`/`BreakFooter` sections to indicate which level they belong to. When there is only one break level, the `BreakLevel` property can be omitted.

Example with 4 break levels (only level 2 has visible header/footer):
```
        Begin BreakLevel
            ControlSource ="OrderDate"
        End
        Begin BreakLevel
            ControlSource ="Region"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Region"
        End
        Begin BreakLevel
            ControlSource ="ArrivalTime"
        End
        ...
        Begin BreakHeader
            BreakLevel =2
            Name ="GroupHeader0"
            ...
        End
        ...
        Begin BreakFooter
            BreakLevel =2
            Name ="GroupFooter1"
            ...
        End
```

## Section Width vs Page Width

A report has a **single** `Width` property, on the `Begin Report` block — it is the width of every section. If it does not fit the printable area, Access reports:

> The section width is greater than the page width…

and prints a blank page after every real page. The constraint, in twips (1 inch = 1440 twips):

```
Report Width + LeftMargin + RightMargin  <=  PaperWidth
```

Margins come from the report's print-settings `.json` (`Items.Margins.LeftMargin` / `.RightMargin`, in **inches** — multiply by 1440); paper size and orientation come from `Items.Printer`.

| Paper | Portrait | Landscape |
|---|---|---|
| A3 | 16838 | 23811 |
| A4 | 11906 | 16838 |
| A5 | 8391 | 11906 |
| Letter | 12240 | 15840 |
| Legal | 12240 | 20160 |

Ignore `Items.Margins.Width` in the `.json` — that is the multi-column item size, not the section width. It is unused when `DefaultSize` is `true` and routinely disagrees with the report `Width`.

### Choosing the value

`Width` must be at least the rightmost control edge, so the smallest legal value is:

```
max(Left + Width) across every control in every section
```

- A control that **omits `Left` has `Left = 0`**. Full-width banner labels in a report header are typically the widest control on the report and are easy to miss for exactly this reason.
- Include controls nested in a child `Begin ... End` block (such as a text box's attached label) — their `Left` is absolute, not relative to the parent.
- `LayoutCached*` properties cache the control's right and bottom edges but go stale. Compute from `Left` and `Width`; never read the cache.

Set `Width` to that maximum. It sits flush against the widest control and leaves no slack to spill onto a second page. If the maximum exceeds the printable width, narrow or reposition controls — or change orientation, paper size, or margins. Raising `Width` past the printable width is never the fix.

## Label Property Restrictions

Label controls do **not** support `BorderStyle` or `OldBorderStyle`. Setting either causes `This property does not apply to this control` on import. To visually distinguish labels (e.g., column headers in a report page header), use background fill instead:

- Set `BackStyle =1` (opaque) — required for `BackColor` to be visible
- Set `BackColor =<color>` to apply a fill color

### `BackStyle` values

| Value | Meaning | Notes |
|-------|---------|-------|
| `0` | Transparent — `BackColor` ignored, section shows through | Default for OptionGroup |
| `1` | Normal — interior filled with `BackColor` | Default for all other controls |

A control inherits `BackStyle` from the report's default control style. If that default sets `BackStyle =0`, every new label is transparent unless overridden.

Wrong (causes import error):
```
Begin Label
    Name ="Header_Label"
    Caption ="Column Header"
    BorderStyle =1            <- ERROR: not applicable to Label
End
```

Correct:
```
Begin Label
    BackStyle =1
    Name ="Header_Label"
    Caption ="Column Header"
    BackColor =14277081       <- light grey background fill
End
```

## Common Structural Errors

- Using `Begin Sorting` → `Expected: 'Begin'. Found: Sorting.` (`Sorting` is not valid syntax; use `BreakLevel` instead)
- Using `Begin GroupHeader0` or `Begin GroupFooter0` → `Expected: Object Type Name. Found: GroupHeader0.` (use `BreakHeader`/`BreakFooter` instead)
- Using `Begin ReportHeader` or `Begin ReportFooter` → use `Begin FormHeader`/`Begin FormFooter` with the appropriate `Name` property instead
- Placing `Begin Label` or sections directly inside `Begin Report` without the container `Begin` → `Expected: 'Begin'. Found: Label.`
