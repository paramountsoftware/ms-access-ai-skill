# Images Reference

Guide for adding, modifying, and replacing images in Access VCS form and report `.bas` files.

## Image Control Types

Access forms and reports can display images via two control types:

| Control | Block keyword | Use case |
|---|---|---|
| `Image` | `Begin Image` | Preferred. Displays a bitmap directly. Lightweight, no OLE overhead. |
| `UnboundObjectFrame` | `Begin UnboundObjectFrame` | Legacy. Embeds an OLE object (e.g., `Word.Document.8`). Much larger file size (~12,800 hex lines vs ~1,500 for a typical image). |

**Always use `Image` controls for static images.** `UnboundObjectFrame` OLE objects should be converted to `Image` controls when encountered — see [Converting OLE to Image](#converting-unboundobjectframe-ole-to-image).

## Image Control Structure

### In Forms

```
Begin Image
    SizeMode =3
    Left =4705
    Top =56
    Width =1756
    Height =1186
    Name ="Image1"
    Picture ="logo.bmp"
    ImageData = Begin
        0x424d1ec00000000000003600000028000000b700000059000000010018000000 ,
        0x0000e8bf0000c40e0000c40e00000000000000000000ffffffffffffffffffff ,
        ...hex lines (32 bytes each, trailing ` ,` except last line)...
        0xffffffffffffffffffffffffffffffffffffffffffffffffffffff000000
    End

    LayoutCachedLeft =4705
    LayoutCachedTop =56
    LayoutCachedWidth =6461
    LayoutCachedHeight =1242
End
```

### In Reports

Report Image controls are structurally identical but may include additional cosmetic properties inherited from the default style block:

```
Begin Image
    OldBorderStyle =0
    Width =1590
    Height =990
    BackColor =16777215
    BorderColor =0
    Name ="Image4"
    Picture ="logo.bmp"
    GridlineColor =0
    ImageData = Begin
        0x424d...
    End

    LayoutCachedWidth =1590
    LayoutCachedHeight =990
    TabIndex =3
End
```

### Key Differences: Forms vs Reports

| Property | Forms | Reports |
|---|---|---|
| `OverlapFlags` | **NOT valid** on Image controls — causes `This property does not apply to this control` on import | **NOT valid** — same restriction |
| `TabStop` | Valid (seen on form Image controls) | Valid |
| `TabIndex` | Valid but often omitted (implicit 0) | Valid, typically present |
| `OldBorderStyle`, `BackColor`, `BorderColor`, `GridlineColor` | Uncommon, inherited from default styles | Common, inherited from default styles |
| `SizeMode` | Typically `=1` (Stretch) or `=3` (Zoom) | Typically `=3` (Zoom) in default styles |

## Invalid Properties on Image Controls

These properties exist on other control types but cause import errors on `Image`:

| Property | Error | Notes |
|---|---|---|
| `OverlapFlags` | `This property does not apply to this control` | Valid on Label, CommandButton, Rectangle, UnboundObjectFrame — but **not** Image |
| `SpecialEffect` | `This property does not apply to this control` | Carried over from UnboundObjectFrame — not valid on Image |
| `Class` | N/A | OLE-specific, not applicable |
| `OLEClass` | N/A | OLE-specific, not applicable |

**When converting from UnboundObjectFrame to Image, strip these properties.**

## ImageData Format

`ImageData` contains a hex-encoded image file. The format is:

```
ImageData = Begin
    0x<64 hex chars = 32 bytes> ,
    0x<64 hex chars = 32 bytes> ,
    ...
    0x<remaining hex chars>
End
```

Rules:
- Every line except the last ends with ` ,` (space-comma)
- The last line has **no** trailing comma
- Indent hex lines one level deeper than the `ImageData = Begin` line

### Supported Image Formats

Access `LoadFromText` accepts multiple image formats in `ImageData`:

| Format | Magic bytes | `Picture` value | Notes |
|--------|-------------|-----------------|-------|
| **PNG** | `0x89504e47` | `"name.png"` | **Preferred.** Smallest file size, supports transparency. |
| **BMP** | `0x424d` | `"name.bmp"` | Legacy. 24-bit uncompressed. Much larger than PNG (~25x for typical logos). |
| **JPEG** | `0xffd8ff` | `"name.jpg"` | Good for photos. Lossy compression. |

**Always prefer PNG** for new images — it produces dramatically smaller `.bas` files (e.g., 108 hex lines vs ~2,700 for an equivalent BMP).

### Embedding a PNG/JPG Directly

Read the image file and hex-encode the raw bytes into lines of 32 bytes (64 hex chars) each. No format conversion is needed:

```python
def image_to_hex_lines(image_path, indent):
    with open(image_path, 'rb') as f:
        raw_bytes = f.read()
    hex_str = raw_bytes.hex()
    lines = []
    for i in range(0, len(hex_str), 64):
        chunk = hex_str[i:i + 64]
        if i + 64 < len(hex_str):
            lines.append(f"{indent}0x{chunk} ,")
        else:
            lines.append(f"{indent}0x{chunk}")
    return lines
```

Set the `Picture` property to match the source format (e.g., `Picture ="logo.png"`).

### Converting to BMP (Legacy)

Only needed if targeting very old Access versions. Use Python with Pillow:

```python
from PIL import Image
import io

def png_to_bmp_bytes(png_path):
    img = Image.open(png_path)
    if img.mode == "RGBA":
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        img = bg
    elif img.mode != "RGB":
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="BMP")
    return buf.getvalue()
```

## Adding a New Image Control

1. **Hex-encode** the source image (PNG preferred) — see above.
2. **Insert** a `Begin Image...End` block at the desired position within the section's control container (`Begin...End` inside the section).
3. **Set required properties:**
   - `SizeMode` — `1` (Stretch) or `3` (Zoom)
   - `Left`, `Top`, `Width`, `Height` — position and size in twips (1 inch = 1440 twips)
   - `Name` — unique control name (e.g., `"Image1"`)
   - `Picture` — filename reference (e.g., `"logo.bmp"`) — does not need to exist on disk
   - `ImageData = Begin...End` — the hex-encoded BMP
4. **Add LayoutCached values** after the `ImageData` End:
   ```
   LayoutCachedLeft =<Left>
   LayoutCachedTop =<Top>
   LayoutCachedWidth =<Left + Width>
   LayoutCachedHeight =<Top + Height>
   ```
5. **Do NOT add** `OverlapFlags`, `SpecialEffect`, or other properties from the invalid list above.
6. **TabIndex:** Only assign if the Image needs to participate in tab order. Image controls can't receive focus, so TabIndex is optional. If assigned, ensure contiguity within the section (see forms/reports references).

## Modifying an Existing Image (Replacing ImageData)

To swap the image displayed by an existing Image control:

1. **Replace** the `Picture` property value with the new filename (use `.png` extension for PNG sources).
2. **Replace** the entire `ImageData = Begin...End` block with the new hex-encoded image data (PNG preferred).
3. **Keep all other properties unchanged** — Name, position, TabIndex, border settings, etc.

Do **not** manually edit individual hex lines within `ImageData`. Always replace the entire block.

## Converting UnboundObjectFrame (OLE) to Image

OLE objects embedded as `UnboundObjectFrame` controls should be converted to `Image` controls. The OLE data cannot be safely edited in place.

### Steps

1. **Identify the OLE control** — look for `Begin UnboundObjectFrame` blocks containing `OleData = Begin...End` and `Class ="Word.Document.8"` (or similar).
2. **Extract positioning properties** from the OLE block: `Left`, `Top`, `Width`, `Height`, `TabIndex`, `Name`, `Visible`.
3. **Discard OLE-specific properties:** `SpecialEffect`, `BackStyle`, `OldBorderStyle`, `OverlapFlags`, `OleData`, `Class`, `OLEClass`.
4. **Build a replacement** `Begin Image...End` block with the new BMP data (see [Adding a New Image Control](#adding-a-new-image-control)).
5. **Replace** the entire `Begin UnboundObjectFrame...End` block (including trailing `LayoutCached*` lines and the closing `End`) with the new `Begin Image...End` block.
6. **Verify** no VBA event handlers reference the control — OLE controls can have events; Image controls cannot fire `Updated`, `Enter`, `Exit` events. (Logo images typically have no events wired.)

### Parsing Caution: Depth Tracking

When programmatically finding the boundaries of a control block, track `Begin`/`End` depth correctly. The binary text format has **two** forms of `Begin`:

- `Begin ControlType` — standard block opener (e.g., `Begin Image`, `Begin Label`)
- `Property = Begin` — inline data block (e.g., `OleData = Begin`, `ImageData = Begin`, `RecSrcDt = Begin`, `OnClickEmMacro = Begin`). Listed here for *depth tracking* only — these share a syntax, not a meaning. `OleData`/`ImageData` are hex payloads; `RecSrcDt` is a single 8-byte timestamp (see [Forms](forms.md) rule 3); `OnClickEmMacro` is an embedded macro.
- `Begin` — bare container opener (section control containers)

All three forms have matching `End` lines. A depth tracker must account for all three:

```python
def is_begin_line(stripped_line):
    return (stripped_line == "Begin" or
            stripped_line.startswith("Begin ") or
            stripped_line.endswith("= Begin"))
```

Failing to track `Property = Begin` patterns will cause the parser to exit the block too early when it encounters the property's `End`, leaving orphaned lines (like `Class`, `LayoutCached*`, and the real closing `End`).

## Default Style Block

The default control styles area (inside the form/report container, before sections) may include a `Begin Image...End` block defining baseline properties for new Image controls. This is **not** a visible control — do not add `Name`, `Picture`, or `ImageData` to it.

```
Begin Image
    BackStyle =0
    OldBorderStyle =0
    SizeMode =3
    PictureAlignment =2
    Width =1701
    Height =1701
    LeftPadding =30
    TopPadding =30
    RightPadding =30
    BottomPadding =30
    GridlineStyleLeft =0
    GridlineStyleTop =0
    GridlineStyleRight =0
    GridlineStyleBottom =0
    GridlineWidthLeft =1
    GridlineWidthTop =1
    GridlineWidthRight =1
    GridlineWidthBottom =1
End
```

## Deleting an Image Control

1. Remove the entire `Begin Image...End` block including trailing `LayoutCached*` lines.
2. **Renumber TabIndex** values in the section to fill the gap (see [forms.md](forms.md) Rule 1).
3. **Recalculate form/report Width** if the deleted control was at the rightmost edge.
