# Forms Reference

> See also [Conditional Formatting](conditional-formatting.md) for format conditions on form controls.

## Critical Rules

These rules apply when editing any form `.bas` file. Violations cause silent corruption or `Error 2128` on import.

1. **TabIndex values MUST be contiguous within each section (0, 1, 2, 3...).** Gaps cause `Error 2128` on import. When removing a control, renumber all subsequent TabIndex values in that section downward to fill the gap.

2. **VBA event procedures MUST be wired up in the `.bas` file.** Adding a `Private Sub Form_Load()` (or any event) to the `.cls` code-behind is **not sufficient** — Access will silently ignore it. You must also add the corresponding event property in the `.bas` layout file. The property goes on the **object that owns the event**:
   - **Form-level events** (`Form_Load`, `Form_Current`, `Form_Open`, etc.) → property on the form object itself (top-level, before sections): `OnLoad ="[Event Procedure]"`, `OnCurrent ="[Event Procedure]"`, etc.
   - **Control events** (`cmdOpen_Click`, `txtName_AfterUpdate`, etc.) → property on the control block: `OnClick ="[Event Procedure]"`, `AfterUpdate ="[Event Procedure]"`, etc.
   - **Section events** (`Detail_Format`, `FormHeader_Click`, etc.) → property on the section block: `OnFormat ="[Event Procedure]"`, `OnClick ="[Event Procedure]"`, etc.

   The property name maps from the VBA event name by dropping the object prefix and using the `On` + PascalCase form. Common mappings:

   | VBA Event Sub | `.bas` Property |
   |---|---|
   | `Form_Load` | `OnLoad ="[Event Procedure]"` |
   | `Form_Current` | `OnCurrent ="[Event Procedure]"` |
   | `Form_Open` | `OnOpen ="[Event Procedure]"` |
   | `Form_Close` | `OnClose ="[Event Procedure]"` |
   | `Form_BeforeUpdate` | `BeforeUpdate ="[Event Procedure]"` |
   | `Form_AfterUpdate` | `AfterUpdate ="[Event Procedure]"` |
   | `cmdButton_Click` | `OnClick ="[Event Procedure]"` (on the control) |
   | `txtField_AfterUpdate` | `AfterUpdate ="[Event Procedure]"` (on the control) |

   **Without this property, the VBA code compiles but never executes.**

3. **Never hand-edit binary data blocks.** Leave `ImageData = Begin...End`, `OleData = Begin...End`, and other hex payload blocks untouched — a single byte error corrupts them. For `ConditionalFormat` and `ConditionalFormat14`, see [Conditional Formatting](conditional-formatting.md) — these must be removed entirely and replaced with VBA code.

   **Exception — `RecSrcDt`.** This is a single 8-byte little-endian IEEE-754 double holding an OLE Automation Date ("Record Source Date"): the moment Access last saved that object's record-source binding. It is a per-object cache stamp, not a checksum of `RecordSource`, and Access re-stamps it on the next design save. It is **optional** — objects with a `RecordSource` and no `RecSrcDt` import cleanly (verified against existing objects that carry a `RecordSource` with no `RecSrcDt` block). When hand-authoring a new form/report, either omit the block or leave a copied one in place; both are safe. Never invent a value, and never set one in the future. Copied objects legitimately share a token — Access does the same on copy-paste, so a family of forms cloned from one original will all carry that original's stamp.

4. **Preserve Begin/End nesting exactly.** Every `Begin` must have a matching `End`. Mismatched nesting causes import failure.

5. **Control order matters.** Controls are rendered in file order (affects z-order). Don't reorder controls unless intentionally changing layering.

6. **Property values use specific formats:**
   - Numeric: `Width =1701` (units are twips: 1 inch = 1440 twips)
   - String: `Name ="Detail"` (always double-quoted)
   - Boolean flags: presence of property name implies true (e.g., `NotDefault`)
   - **String escaping uses backslash (`\"`), NOT VBA-style doubled quotes (`""`).**  This is the binary text format's own convention and differs from VBA code in `.cls` files. Example — a `DLookUp` ControlSource:
     ```
     ControlSource ="=DLookUp(\"Description\",\"vwMyView\",\"Id=\" & [Id])"
     ```
     Using `""` instead of `\"` will cause the value to be parsed incorrectly on import, corrupting the expression.

7. **When removing a control block,** remove everything from `Begin ControlType` through its closing `End`, including any trailing `LayoutCached*` properties and `PictureCaptionArrangement` that appear between the inner `End` (closing ImageData) and the outer `End` (closing the control). A complete control block looks like:

   ```
   Begin CommandButton              <- Start of control
       Name ="btnExample"
       Caption ="Click Me"
       TabIndex =2
       ImageData = Begin            <- Inner Begin/End for binary data
           0x...
       End                          <- Closes ImageData

       LayoutCachedLeft =226        <- Layout cache (part of this control)
       LayoutCachedTop =113
       LayoutCachedWidth =802
       LayoutCachedHeight =689
       PictureCaptionArrangement =5
   End                              <- Closes CommandButton
   ```

8. **Recalculate form width when adding or modifying controls.** See [Width Recalculation](#width-recalculation) below for the full formula and subform chrome table.

9. **Subform default style blocks have restricted properties.** See [Subform Default Style Restrictions](#subform-default-style-restrictions) below for valid/invalid property lists.

## Form `.bas` Skeleton

Forms use Access's binary text format. **Every `.bas` file must start with a version header** before the `Begin Form` block.

The overall structure has exactly **one** `Begin...End` container block inside `Begin Form`. This container holds **both** the default control styles **and** all sections. Nothing except form-level properties may appear outside this container.

```
Version =20
 VersionRequired =20
Begin Form
    RecordSource ="..."             <- Form-level properties go here
    Caption ="..."
    Width =...
    Begin                           <- ONE container for everything below
        Begin Label                 <- Default control styles (Label, TextBox, CommandButton, etc.)
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin FormHeader            <- Sections follow default styles, INSIDE the same container
            Height =500
            Name ="FormHeader"
            Begin                   <- Section controls container
                Begin Label
                    Name ="lblTitle"
                    ...
                End
            End
        End
        Begin Section               <- Detail section
            Name ="Detail"
            Begin
                ...controls...
            End
        End
        Begin FormFooter
            ...
        End
    End                             <- Closes the ONE container
End                                 <- Closes Begin Form
CodeBehindForm                      <- Required when SplitLayoutFromVBA=true and .cls exists
' See "fFormName.cls"
```

### Form Sections

Forms use these section keywords: `FormHeader`, `Section` (Detail), `FormFooter`, `PageHeader`, `PageFooter`.

### Default Control Styles Area

The default control style blocks (Label, TextBox, CommandButton, etc.) define baseline properties for new controls. They appear at the start of the container, before any sections.

### `CodeBehindForm` Trailer Line

When `SplitLayoutFromVBA=true` and a `.cls` file exists, the `.bas` file must end with `CodeBehindForm` after the closing `End` of `Begin Form`.

### Common Structural Errors

- Placing `Begin Label` or sections directly inside `Begin Form` without the container `Begin` → `Expected: 'Begin'. Found: Label.`
- Wrapping default styles in their own separate `Begin...End` then having sections outside it → `Expected: 'End'. Found: Begin.`

**Key rule:** After the form-level properties, there is exactly ONE `Begin...End` container block. Default control styles and ALL sections live inside it.

## ComboBox Controls

### `RowSourceType ="Value List"` requires `RowSourceTypeInt =1`

`RowSourceType` is only the human-readable label — **`RowSourceTypeInt` is what Access
reads.** When it is absent it defaults to `0` = Table/Query, so a control written as:

```
Begin ComboBox
    Name ="cmbStatus"
    RowSourceType ="Value List"      <- looks right
    RowSource ="Open;Closed"
End
```

imports without complaint and then **throws at runtime** — Access tries to execute
`Open;Closed` as SQL when the form opens or the list drops down. The Access property sheet
can still show "Value List", so the failure looks unrelated to the row source.

Correct — both properties, with the `Int` form near the **top** of the block:

```
Begin ComboBox
    LimitToList = NotDefault        <- Yes: reject values not in the list
    RowSourceTypeInt =1             <- 1 = Value List (REQUIRED)
    OverlapFlags =85
    Name ="cmbStatus"
    RowSourceType ="Value List"
    RowSource ="Open;Closed"
    AllowValueListEdits =0          <- see below
End
```

| `RowSourceTypeInt` | Meaning |
|---|---|
| `0` (or absent) | Table/Query |
| `1` | Value List |
| `2` | Field List |

**Also override `AllowValueListEdits =0`** whenever the form's default ComboBox style block
sets it to `1`, so users cannot edit the hardcoded list at runtime.

**Item quoting.** `RowSource ="Low;Medium;High"` and
`RowSource ="\"Low\";\"Medium\";\"High\""` both work; quote items that could otherwise be
read as a number or an expression.

**Audit** — value-list controls missing the integer property:

```bash
for f in forms/*.bas; do
  n=$(grep -c 'RowSourceType ="Value List"' "$f")
  m=$(grep -c 'RowSourceTypeInt =1' "$f")
  [ "$n" -gt "$m" ] && echo "$f: $n value lists, $m tagged"
done
```

### BoundColumn

`BoundColumn` in `.bas` files is **0-based** — this differs from the Access UI and VBA, which display it as **1-based**. A `.bas` value of `0` corresponds to column 1 in the UI. This offset is a common source of bugs.

**Example:** A ComboBox with `RowSource` returning `ID, CustomerName`:
```
 ColumnCount =2
 ColumnWidths ="0;3000"
BoundColumn =0
```
- Column 0 = `ID` (hidden by `ColumnWidths ="0;..."`)
- Column 1 = `CustomerName` (displayed, width 3000 twips)
- `BoundColumn =0` → stores the `ID` value in `ControlSource`

**Common mistake:** Setting `BoundColumn =1` when the intent is to store the ID. In the `.bas` file, `1` means the **second** column (the display name), not the first. This stores the display name instead of the foreign key, causing data integrity issues. **Always verify that `BoundColumn` points to the primary key column (usually column 0 in `.bas`), not the display column.**

**Rule:** When using a multi-column ComboBox as a foreign key lookup:
1. The first column should be the ID/key field
2. Set `ColumnWidths` to hide the ID column: `"0;3000"` (zero width hides it)
3. Set `BoundColumn =0` to store the ID value (0-based in `.bas` files)
4. Set `ControlSource` to the foreign key field name

## Width Recalculation

After adding, removing, or resizing controls in a form or report, recalculate the form/report `Width` property to equal the rightmost edge of any control (`Left + Width`). Update the `LayoutCachedWidth` of any affected controls and their associated header/footer labels to match. **If the form is used as a subform**, also update the parent/master form:
   - Find the `Begin Subform` block in the parent form that references this form via `SourceObject`
   - Calculate the **subform control width** = child form `Width` + UI chrome (see below)
   - Update the subform control's `Width` and `LayoutCachedWidth` to the calculated value
   - If the calculated value exceeds the parent form's `Width`, increase the parent form's `Width` to fit

**How to find the parent form:** Search all `.bas` files in `forms/` for `SourceObject ="Form.<formname>"` to locate the master form containing the subform control.

**Subform UI chrome:** The child form's `Width` property is only the content area. The subform control in the parent must also account for UI chrome that sits outside the content area. Check the child form's properties and add:

| Chrome Element | Condition (on child form) | Approximate Width |
|---|---|---|
| Record selector | `RecordSelectors = NotDefault` is **absent** (default = visible) | ~340 twips |
| Vertical scrollbar | `ScrollBars` is absent (default = 3/Both) or set to 2 or 3 | ~300 twips |
| Borders | Default border style | ~60 twips (30 per side) |

Formula: `subform control Width = child form Width + record_selector + scrollbar + borders`

**Worked example** — after shrinking `txtField` from 4000 to 3000 twips at `Left =12000`:
- New right edge = 12000 + 3000 = 15000
- Set child form `Width =15000`
- Set `txtField` `LayoutCachedWidth =15000`
- Set associated header label `Width =3000`, `LayoutCachedWidth =15000`
- Child form has record selectors (on) + scrollbar (default/both) + borders: 340 + 300 + 60 = 700
- In parent form: set subform control `Width =15700`, `LayoutCachedWidth =15700`
- In parent form: set form-level `Width =15700` (if new value exceeds current width)

## Subform Default Style Restrictions

In the default control styles area (inside the container), the `Begin Subform...End` block must only contain layout/border properties. **Do NOT include font properties** — they don't apply to Subform controls and cause `This property does not apply to this control` on import.

Valid Subform default style properties: `BorderLineStyle`, `Width`, `BorderColor`, `Height`, `Left`, `Top`.

Invalid (will cause import error): `FontSize`, `FontName`, `FontWeight`, `ForeColor`, `FontItalic`, `FontUnderline`.

Correct:
```
Begin Subform
    BorderLineStyle =0
    Width =1701
    BorderColor =12632256
End
```

Wrong (causes error):
```
Begin Subform
    BorderLineStyle =0
    Width =1701
    FontSize =11            <- ERROR: not applicable to Subform
    BorderColor =12632256
    FontName ="Calibri"     <- ERROR: not applicable to Subform
End
```

## Label Property Restrictions

Label controls do **not** support `BorderStyle` or `OldBorderStyle`. Setting either causes `This property does not apply to this control` on import. To visually distinguish labels (e.g., column headers), use background fill instead:

- Set `BackStyle =1` (opaque) — required for `BackColor` to be visible
- Set `BackColor =<color>` to apply a fill color

### `BackStyle` values

| Value | Meaning | Notes |
|-------|---------|-------|
| `0` | Transparent — `BackColor` ignored, form/section shows through | Default for OptionGroup |
| `1` | Normal — interior filled with `BackColor` | Default for all other controls |

A control inherits `BackStyle` from the form's default control style (the `Begin Label` / `Begin TextBox` block in the form container). If that default sets `BackStyle =0`, every new label on the form is transparent unless overridden.

Wrong (causes import error):
```
Begin Label
    Name ="Header_Label"
    Caption ="Column Header"
    BorderStyle =1            <- ERROR: not applicable to Label
    BorderColor =0            <- meaningless without BorderStyle
End
```

Correct (use background fill instead):
```
Begin Label
    BackStyle =1
    Name ="Header_Label"
    Caption ="Column Header"
    BackColor =14277081       <- light grey background fill
End
```

## List/Detail Form Navigation Pattern

A common Access pattern is a list form (continuous/datasheet) that opens a detail form for add/edit, then needs to refresh when the user returns. The list form's data must be requeried to show changes made in the detail form.

### Pattern: Requery on Return from Detail Form

**List form (`.cls`)** — opens the detail form for viewing/editing:
```vba
Private Sub cmdOpen_Click()
    If Not IsNull(Me.txtID) Then
        DoCmd.OpenForm "fCustomerDetail", , , "ID = " & Me.txtID
    End If
End Sub

Private Sub cmdAdd_Click()
    DoCmd.OpenForm "fCustomerDetail", , , , acFormAdd
End Sub
```

**Detail form (`.cls`)** — saves/discards and requeries the calling list form:
```vba
Private Sub cmdSaveClose_Click()
    If Me.Dirty Then Me.Dirty = False
    RequeryCallingForm
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdDiscardClose_Click()
    Me.Undo
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub RequeryCallingForm()
    ' Requery the list form so it reflects changes
    Dim frm As Form
    For Each frm In Forms
        If frm.Name <> Me.Name Then
            frm.Requery
        End If
    Next frm
End Sub
```

Alternatively, use a **shared module** function for reusable save/close logic:
```vba
' In modNavigation.bas:
Public Sub SaveAndCloseForm()
    Dim frmName As String
    frmName = Screen.ActiveForm.Name
    With Screen.ActiveForm
        If .Dirty Then .Dirty = False
    End With
    ' Requery all other open forms
    Dim frm As Form
    For Each frm In Forms
        If frm.Name <> frmName Then
            frm.Requery
        End If
    Next frm
    DoCmd.Close acForm, frmName
End Sub
```

### Wiring Up Events in `.bas`

Remember that the `.bas` file must have the event properties wired:
```
OnClick ="[Event Procedure]"
```
on each button control (see [Critical Rule 2](#critical-rules) above). Without this, the VBA event handlers will never fire.

### Key Points
- **Always requery the list form** after saving in the detail form, otherwise the list shows stale data
- The requery should happen **before** closing the detail form (while the calling form is still accessible)
- `Me.Dirty = False` forces a save of the current record before closing
- `Me.Undo` discards unsaved changes before closing
- If the list form uses conditional formatting via VBA (`FormatConditions`), re-apply it after requery (see [Conditional Formatting](conditional-formatting.md))
