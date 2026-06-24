# sma_ema

A LibreOffice Calc extension for calculating **simple** and **exponential
moving averages** (SMA / EMA).

It provides one **cell function** you can type into a worksheet:

- `MOVAVG(data, period, [type])` — moving average; `type` is `"s"` (or
  omitted) for a simple moving average, `"e"` for an exponential one.

There are **two implementations** of the same `MOVAVG` function — pick one
(don't install both, or the duplicate definitions collide):

| Implementation | Autocomplete / Function Wizard | Setup |
|----------------|-------------------------------|-------|
| **Basic macro** (`MovingAverages.bas`) | ❌ no — you type the whole name | Paste into the Basic IDE, or install `sma_ema.oxt`. Runs on every LibreOffice with no dependencies. |
| **Python Add-In** (`addin/`) | ✅ yes — autocompletes and shows in the `fx` wizard | Install `sma_ema_addin.oxt`. Needs LibreOffice's Python scripting (bundled on Windows/macOS). |

---

## Installing the extension

1. Download / build `sma_ema.oxt` (see [Building](#building-the-oxt) below).
2. In LibreOffice: **Tools → Extension Manager… → Add…**, choose
   `sma_ema.oxt`, and accept.
3. Restart LibreOffice.

The macros are installed as a Basic library named **SmaEma**, available under
**Tools → Macros**.

---

## Using the cell functions

`MOVAVG` is an ordinary **fill-down function** — it returns the moving average
value at the *end* of the range you give it. You write one formula and drag it
down the column, just like `SUM` or `AVERAGE`.

1. In the first output cell (say `B2`, next to data in column `A`), type the
   formula with the **start of the range anchored** with `$`:
   - `=MOVAVG(A$2:A2;10)` — simple moving average (default)
   - `=MOVAVG(A$2:A2;10;"e")` — exponential moving average
2. Press **Enter**.
3. Select that cell and **drag the fill handle down** the column (or copy it
   down). On each row the range grows — `A$2:A3`, `A$2:A4`, … — so each cell
   shows the moving average up to that row.

Rows that don't yet have `period` values behind them come back **blank**.

> Anchor the start (`A$2`) so the EMA accumulates correctly from the first
> value. Use `;` or `,` as the argument separator depending on your locale.

| Argument | Meaning |
|----------|---------|
| `data` | The input range, e.g. `A$2:A2`. |
| `period` | Window size — a whole number ≥ 1. |
| `type` | `"s"` or omitted = **simple** (trailing average of the last `period` values); `"e"` = **exponential** (`alpha = 2 / (period + 1)`, seeded with the simple average of the first `period` values). |

Non-numeric / blank cells in the input are treated as `0`.

---

## The Python Add-In (autocomplete + Function Wizard)

The Basic function works but never autocompletes — that's a hard limit of
Basic user functions in LibreOffice. For a `MOVAVG` that **completes as you
type** and appears in the `fx` **Function Wizard** with argument help, install
the Python Calc Add-In instead. It computes exactly the same values; only the
integration differs.

**Install**

1. Build `sma_ema_addin.oxt` (see [Building the add-in](#building-the-add-in)).
2. **Tools → Extensions… → Add…**, choose `sma_ema_addin.oxt`, accept.
3. Restart LibreOffice.
4. If you also have the **Basic** `MOVAVG` (in a Basic library or the other
   extension), remove it — two definitions of `MOVAVG` collide.

Then `=MOV…` autocompletes to `MOVAVG`, and the Function Wizard lists it under
the **Add-In** category. Usage (range, period, type) is identical to above.

**Requirements:** LibreOffice's Python scripting, which is bundled on Windows
and macOS. Some Linux packages need `libreoffice-script-provider-python`
installed separately. No third-party Python libraries are used.

---

## Project layout

| File | Purpose |
|------|---------|
| `MovingAverages.bas` | The Basic `MOVAVG` cell function (simple + exponential). |
| `build.ps1` | Packages the Basic module into `sma_ema.oxt`. |
| `addin/idl/MovingAverages.idl` | UNO interface declaring the add-in's `MOVAVG` function. |
| `addin/MovingAveragesAddIn.py` | The Python Add-In component (same math, in Python). |
| `addin/CalcAddIns.xcu` | Function Wizard metadata (name, argument descriptions). |
| `addin/description.xml`, `addin/META-INF/manifest.xml` | Add-In extension metadata. |
| `build_addin.ps1` | Compiles the IDL and packages the add-in into `sma_ema_addin.oxt`. |

The `.bas` and `addin/` sources are the source of truth; the `.oxt` files and
`addin/MovingAverages.rdb` are generated.

---

## Building the .oxt

The extension is built from the `.bas` source by a PowerShell script:

```powershell
pwsh ./build.ps1
```

This produces `sma_ema.oxt` in the project root. Internally it converts the
`.bas` module into the LibreOffice `.xba` format, adds the extension metadata
(`description.xml`, `META-INF/manifest.xml`, library index), and zips it.

### Building the add-in

The Python Add-In is built by a separate script that needs the **LibreOffice
SDK** installed (for the `unoidl-write` type-library compiler):

```powershell
pwsh ./build_addin.ps1
```

It compiles `addin/idl/MovingAverages.idl` into `addin/MovingAverages.rdb`,
then zips the Python component, type library, `CalcAddIns.xcu`, and metadata
into `sma_ema_addin.oxt`. If LibreOffice isn't at the default path, pass
`-LoRoot "C:\path\to\LibreOffice"`.

---

## Development without building

You can also use the macros without packaging:

1. **Tools → Macros → Edit Macros** to open the Basic IDE.
2. Create a module and paste the contents of `MovingAverages.bas`.
3. Use the `MOVAVG` function as above.
