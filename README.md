# sma_ema

A LibreOffice Calc extension for calculating **simple** and **exponential
moving averages** (SMA / EMA).

It provides two **cell functions** you can type into a worksheet:

- `SMA(data, period)` — simple moving average
- `EMA(data, period)` — exponential moving average

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

`SMA` and `EMA` are **array functions** — they return a whole column at once.

1. Select an output range the same height as your data.
2. Type the formula, e.g. `=SMA(A2:A100;10)` or `=EMA(A2:A100;10)`.
3. Confirm with **Ctrl+Shift+Enter**.

Rows that don't yet have enough data (the first `period-1` rows) come back
empty.

| Function | Definition |
|----------|------------|
| `SMA(data, period)` | Trailing average of the last `period` values. |
| `EMA(data, period)` | Smoothing factor `alpha = 2 / (period + 1)`, seeded with the simple average of the first `period` values. |

Non-numeric / blank cells in the input are treated as `0`.

---

## Project layout

| File | Purpose |
|------|---------|
| `MovingAverages.bas` | The `SMA` / `EMA` cell functions plus the shared `ComputeSMA` / `ComputeEMA` math. |
| `build.ps1` | Packages the module into `sma_ema.oxt`. |

The `.bas` file is the source of truth; the `.oxt` is generated from it.

---

## Building the .oxt

The extension is built from the `.bas` source by a PowerShell script:

```powershell
pwsh ./build.ps1
```

This produces `sma_ema.oxt` in the project root. Internally it converts the
`.bas` module into the LibreOffice `.xba` format, adds the extension metadata
(`description.xml`, `META-INF/manifest.xml`, library index), and zips it.

---

## Development without building

You can also use the macros without packaging:

1. **Tools → Macros → Edit Macros** to open the Basic IDE.
2. Create a module and paste the contents of `MovingAverages.bas`.
3. Use the `SMA` / `EMA` functions as above.
