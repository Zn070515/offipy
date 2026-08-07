> [中文](assets.md)

# Asset System (offipy.assets)

offipy v0.14 introduces a unified **Asset System v1**: it abstracts visual material
(icons, textures, native primitives) as `asset://` resources that flow through one
pipeline — **provider resolve → deterministic measurement → placeholder injection →
render placement** — and end up as **editable native PowerPoint objects** (native
freeforms / SVG pictures / native shapes and text boxes) instead of flattened rasters.

- **Deterministic**: the same HTML + the same theme produces a consistent structure
  every time; asset declarations get stable IDs.
- **Editable**: icons become native freeforms; primitives become native shapes + text
  boxes (double-click to edit the text).
- **Provenance**: every render emits an `assets.json` usage manifest recording each
  asset's provider / license / source for compliance auditing.
- **Offline**: icons / textures / primitives are all vendored into the wheel; rendering
  never hits the network.

## Quick start

```html
<div data-asset="asset://ph/icon/check"
     style="position:absolute;left:80px;top:140px;width:120px;height:120px;color:var(--accent)"></div>
<div data-asset="asset://procedural/pattern/topography?seed=42" data-asset-placement="background"
     style="position:absolute;left:0;top:0;width:100%;height:100%"></div>
<div data-asset="asset://primitives/primitive/metric-badge"
     data-asset-param-value="24%" data-asset-param-label="YoY" data-asset-param-delta="+3.2%"
     style="position:absolute;left:100px;top:200px;width:420px;height:220px"></div>
```

```bash
offipy deck make --html deck.html --theme mckinsey --layouts --out deck.pptx
```

## Built-in providers in v0.14

| provider | kind | content | license | first_party |
|----------|------|---------|---------|-------------|
| `ph` | `icon` | 1512 Phosphor icons (vendored) | MIT | no |
| `lu` | `icon` | 1756 Lucide icons (vendored) | ISC | no |
| `procedural` | `pattern` | 8 deterministic textures (wave / blob / dot-grid / square-grid / rings / topography / circuit / gradient-orb) | MIT | yes |
| `primitives` | `primitive` | 8 editable native primitives (quote-mark / section-number / label-pill / metric-badge / timeline-node / process-arrow / device-frame / browser-mockup) | MIT | yes |

`AssetKind` reserves `illustration` / `map` / `flag` kinds for future external
providers (v0.15 milestone); they are not implemented in v0.14.

## URI grammar

```text
asset://<provider>/<kind>/<name>[?k=v&k2=v2...]
```

- `<provider>` / `<kind>` / `<name>` are lowercase, hyphen-separated.
- Query params `k=v` are canonicalized (`_`→`-`, lower-cased, sorted); keys must be
  unique; values are percent-decoded. Hex colors round-trip as `%23RRGGBB`
  (`#` → `%23`).
- `#` fragments and CSS `var(` references are rejected (param values must not
  reference CSS variables).
- Examples:

```text
asset://ph/icon/check
asset://lu/icon/settings
asset://procedural/pattern/topography?seed=42
asset://procedural/pattern/rings?count=4
asset://primitives/primitive/metric-badge
```

### Render modes

| mode | meaning |
|------|---------|
| `freeform_svg` | source SVG parsed into a native PowerPoint freeform (`p:sp` + `a:custGeom`), double-click to edit the path (icons) |
| `svg` | SVG written as an OOXML SVG picture — primary `a:blip` embeds a PNG raster fallback, `asvg:svgBlip` points at the vector (procedural textures; degrades to pure SVG without Playwright) |
| `svg_template` | template SVG materialized via color slots (e.g. `__ACCENT__`) then written as an SVG picture |
| `raster` | bitmap payload (`add_picture`) |
| `native_shape` | native shapes + text boxes + freeforms combined (primitives) |

### Placement

| value | meaning |
|-------|---------|
| `replace` (default) | rendered elements are inserted at the declaration's original DOM slot, replacing the placeholder |
| `decorative` | same slot semantics as replace, tagged as decorative |
| `background` | moved after `grpSpPr` and before all content shapes (bottom layer) |

> **v0.14 constraint**: `background` is only available for vector resources such as
> procedural textures. Declaring `data-asset-placement="background"` on a native
> primitive **fails explicitly**.

## HTML usage

### Declaring an asset

```html
<div data-asset="asset://<uri>"
     data-asset-param-<key>="<value>"
     data-asset-placement="replace|decorative|background"
     style="..."></div>
```

- `data-asset-param-*` passes params per the provider's schema; hex colors are escaped
  with `%23` (e.g. `data-asset-param-accent="%232251ff"`), or pass a theme semantic
  token directly (`accent` / `surface` / `ink` / `muted` / `bg`).
- The asset size comes from the element's measured rect (real browser BCR); the
  rendered result is scaled into that rect.

### Legacy `data-icon` compatibility

The v0.12+ icon container syntax is unchanged; internally it migrates to the `ph` /
`lu` providers:

```html
<svg data-icon="ph:check-circle" viewBox="0 0 256 256" width="72" height="72"></svg>
```

`data-icon="<set>:<name>"` is equivalent to `asset://ph/icon/<name>` (or `lu`); there
is no behavior change in v0.14 and no migration required.

### `data-primitive` sugar

Native primitives have a shorthand:

```html
<div data-primitive="metric-badge" data-asset-param-value="24%"></div>
```

Preprocessing expands this to the canonical `data-asset="asset://primitives/primitive/metric-badge"`
plus an internal ID, fully equivalent to the canonical form.

### Preflight for declaration injection

Asset injection depends on the `measurements.json` produced by visual audit. Therefore:

- `no_visual_audit=True` (CLI `--no-visual-audit`) is **incompatible** with
  `data-asset` / `data-primitive` / `data-icon` declarations and fails fast before
  launching chromium / convert.
- After a visual-audit render, the output directory (e.g. `out_audit/`) contains
  `assets.json`; `no_visual_audit` does not produce an asset manifest.

## Python API

```python
from offipy.assets import get_default_registry

r = get_default_registry()

# search
metas = r.search("grid", kind="pattern")      # filter by kind
print([m.ref for m in metas])

# resolve
asset = r.resolve("asset://procedural/pattern/topography?seed=42")
print(asset.meta.ref, asset.provider_meta.license)

# primitives with required params: use AssetRequest (example uses a primitive with
# optional-only params)
from offipy.assets import AssetRef, AssetRequest

resolved = r.resolve(AssetRequest(AssetRef("primitives", "primitive", "browser-mockup")))
assert resolved.payload.primitive == "browser-mockup"
```

`import offipy.assets` is a pure-stdlib surface (model / uri / registry / license /
color / materialize) and **does not** load python-pptx / Pillow / playwright.

## assets.json usage manifest

After a visual-audit render, `_audit/assets.json` in the output directory records every
asset used in that render:

```json
{
  "schema": 1,
  "assets": [
    {
      "declaration_id": "asset-s01-001",
      "slide_index": 1,
      "request": "asset://primitives/primitive/metric-badge",
      "provider": {"id": "primitives", "license": "MIT", "source_url": "...", "first_party": true},
      "placement": "replace"
    }
  ]
}
```

- `declaration_id`: stable, deterministically incremented (`asset-s<slide>-<seq>`).
- Charts (`data-chart`) are **not** Asset System v1 resources and never appear in the
  manifest.

## Native primitive schema (v0.14)

All primitives share optional `accent` (default theme accent) and `fill`
(primitive-specific default). Params are validated at the provider layer; unknown
params fail explicitly.

| primitive | params | default fill | structure |
|-----------|--------|--------------|-----------|
| `quote-mark` | `text` (required, ≤240) | transparent | quote glyph + text box |
| `section-number` | `number` (required int 0..9999), `label` (optional ≤120) | surface | number + optional label + accent decoration |
| `label-pill` | `text` (required, ≤120) | accent | rounded rectangle + centered text |
| `metric-badge` | `value` (required ≤80), `label` (optional ≤120), `delta` (optional ≤40) | surface | card + value/label/delta |
| `timeline-node` | `label` (optional ≤120), `phase` (optional past/current/future, default current) | — | node dot + label; phase drives style |
| `process-arrow` | `steps` (required comma-separated 2..8 items, each ≤80), `direction` (optional horizontal/vertical, default horizontal) | surface | arrow split into N segments, each text editable |
| `device-frame` | `device` (required phone/tablet/desktop) | surface | native device shell + empty screen |
| `browser-mockup` | `title` (optional ≤120), `url` (optional ≤240) | surface | window card + chrome bar + three dots |

Known limitations (explicit for v0.14, not bugs):

- `process-arrow` `steps` does not support escaped commas; use another primitive if a
  label needs a comma.
- `device-frame` / `browser-mockup` accept **no screenshot / nested asset**: there is no
  `screenshot` / `src` / `image` param; the screen is an empty editable area.
- Primitive text uses a default safe sans font (Arial family); a theme font token is a
  v0.15 direction, and v0.14 introduces no new public font-token model.

## Non-goals / future providers

- v0.15 milestone: external illustration / map / photo providers (`illustration` /
  `map` / `flag`), which need network or large assets; not in v0.14.
- Nested assets (e.g. putting a screenshot inside a device frame) are an explicit
  **non-goal** and fail loudly rather than degrading silently.
