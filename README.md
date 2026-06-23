# sma_ema

A LibreOffice Calc extension for calculating **simple** and **exponential
moving averages** (SMA / EMA).

It provides two things:

1. **Cell functions** you can type into a worksheet:
   - `SMA(data, period)` — simple moving average
   - `EMA(data, period)` — exponential moving average
2. **A dialog box** that lets you pick the input data, the output start cell,
   and one or more window sizes, then writes the result columns for you.

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

## Using the dialog

Run **Tools → Macros → Run Macro… → My Macros / SmaEma →
MovingAverageDialog → ShowMovingAverageDialog** (or bind it to a toolbar
button).

The dialog lets you set:

- **Input data range** — pre-filled from the current selection (e.g. `A2:A100`).
- **Output start cell** — where the first result column begins (e.g. `D2`);
  output is aligned row-for-row with the input.
- **Window sizes** — comma-separated, e.g. `5,10,20`.
- **SMA / EMA** — tick either or both.

On OK it writes one column per measure × window, each with a header label
(`SMA10`, `EMA20`, …) in the row directly above the output start cell. For
example, with both boxes ticked and windows `5,10` you get four columns:
`SMA5`, `EMA5`, `SMA10`, `EMA10`.

Output is written as plain values (it does not recalculate when the source
data changes). Input is treated as a single column.

---

## Project layout

| File | Purpose |
|------|---------|
| `MovingAverages.bas` | The `SMA` / `EMA` cell functions plus the shared `ComputeSMA` / `ComputeEMA` math. |
| `MovingAverageDialog.bas` | `ShowMovingAverageDialog` — the runtime-built dialog and the code that writes results. |
| `build.ps1` | Packages both modules into `sma_ema.oxt`. |

The two `.bas` files are the source of truth; the `.oxt` is generated from
them.

---

## Building the .oxt

The extension is built from the `.bas` source by a PowerShell script:

```powershell
pwsh ./build.ps1
```

This produces `sma_ema.oxt` in the project root. Internally it converts each
`.bas` module into the LibreOffice `.xba` format, adds the extension metadata
(`description.xml`, `META-INF/manifest.xml`, library index), and zips it.

---

## Development without building

You can also use the macros without packaging:

1. **Tools → Macros → Edit Macros** to open the Basic IDE.
2. Create a module and paste the contents of each `.bas` file.
3. Use the functions / run `ShowMovingAverageDialog` as above.
