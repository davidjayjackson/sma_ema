# sma_ema

A LibreOffice Calc extension for calculating **simple** and **exponential
moving averages** (SMA / EMA).

It provides one **cell function** you can type into a worksheet:

- `MOVAVG(data, period, [type])` — moving average; `type` is `"S"` (or
  omitted) for a simple moving average, `"E"` for an exponential one.

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
   - `=MOVAVG(A$2:A2;10;"E")` — exponential moving average
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
| `type` | `"S"` or omitted = **simple** (trailing average of the last `period` values); `"E"` = **exponential** (`alpha = 2 / (period + 1)`, seeded with the simple average of the first `period` values). |

Non-numeric / blank cells in the input are treated as `0`.

---

## Project layout

| File | Purpose |
|------|---------|
| `MovingAverages.bas` | The `MOVAVG` cell function (simple + exponential moving averages). |
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
3. Use the `MOVAVG` function as above.
