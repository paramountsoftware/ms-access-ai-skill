# VBA Reference

> Cross-reference: [Forms Critical Rules](forms.md#critical-rules) apply — especially Rule 2 (event wiring). See also [Conditional Formatting](conditional-formatting.md) for VBA `FormatConditions` patterns.

## VBA Modules (modules/*.bas)

- Start with `Attribute VB_Name = "ModuleName"`
- Use `#If VBA7 Then` / `#Else` / `#End If` for API declarations (required for 64-bit Access)
- `Option Compare Database` is standard

## VBA Code-Behind Files (.cls)

Form and report code-behind files (`.cls`) are created when `SplitLayoutFromVBA=true`. These files must **NOT** include the `VERSION 1.0 CLASS` header block or `Attribute VB_Name`. VCS handles class registration internally during import — including these causes import errors.

**Correct `.cls` format:**
```vba
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOpen_Click()
    DoCmd.OpenForm "fDetail"
End Sub
```

**Wrong — do NOT include:**
```vba
VERSION 1.0 CLASS          <- WRONG: remove this entire block
BEGIN                       <- WRONG
  MultiUse = -1  'True     <- WRONG
END                         <- WRONG
Attribute VB_Name = "Form_fMyForm"  <- WRONG: remove this line
Attribute VB_GlobalNameSpace = False
...
```

The `VERSION 1.0 CLASS` block and `Attribute VB_Name` are standard VBA class file headers used by the VB editor, but VCS strips them on export and does not expect them on import.

## Writing to Hyperlink Columns

An Access hyperlink value is a `#`-delimited triplet: `displaytext#address#subaddress`.
Two rules follow; breaking either produces a link that looks fine in the grid and opens
the wrong thing:

1. **The address must be absolute** (`C:\...` or `\\server\share\...`). Access resolves a
   relative address against the folder the front end runs from, so the same row works from
   one copy of the front end and fails from another.
2. **No `#` anywhere in the display text or the address.** Windows *permits* `#` in
   filenames, which is what makes it dangerous — the extra delimiter shifts the triplet,
   the address segment becomes a bare relative fragment, and that then hits rule 1. One
   reference-style filename (`Invoice#123.pdf`) is enough to do it.

Sanitise generated filenames at the source (strip `\/:*?"<>|` **and** `#`) rather than
escaping at each use. For a file the user picked off disk the name cannot be sanitised —
prompt the user to rename it instead of storing a link you already know is broken.

Own the sanitise / build / read-back helpers in one shared module rather than repeating
them per form, and read an address back with `HyperlinkPart(value, acAddress)` — never
split on `#` by hand.

## Module-Level Declarations

Put `Const` and `Dim` declarations in the declarations section at the top of the module,
above the first procedure. VBA tolerates them between procedures, but they read as stray
statements and hide from anyone scanning the header.
