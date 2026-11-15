# Configuration

Every adjustable constant is at the top of `src/ConstraintBasedProcessor.bas`.
**Source code is the single source of truth** — numbers here reference
names only.

## Five Modules

| Module | Entry | What it does | Input sheet | Output |
|---|---|---|---|---|
| **Clean & Normalize** | `RunModule1()` | Strip suffixes, pad postal codes, backfill missing codes | Input | Input (in-place) |
| **Validate & Guard** | `RunModule3()` | Count-consistency check before split | Input | Input (red-font flags) |
| **Split & Redistribute** | `RunModule2()` | Split slash-separated rows, redistribute weight/volume, absorb drift | Input | Input (rows inserted) |
| **Constraint Convergence** | `RunModule4()` | Iterative dimension uniform-adjustment until aggregate volume in target range | Input | Input (dimensions modified) |
| **Route & Classify** | `RunModule5()` | Regex-match tracking numbers → carrier codes | Input | Input (routing column filled) |
| **Run all** | `RunFullPipeline()` | Init → M1 → M3 → M2 → M4 → M5 | Input + Lookup | Input |

## Sheet Name Constants

| Constant | Default | When to change |
|---|---|---|
| `SHEET_INPUT`  | `"Input"`  | Your primary data sheet |
| `SHEET_LOOKUP` | `"Lookup"` | Your lookup table for code backfill |

## Column Letter Constants

All single Excel column letters.

| Constant | Default | What it is |
|---|---|---|
| `COL_CODE`     | `"E"` | Entity code (suffix-stripped) |
| `COL_PART_NO`  | `"C"` | Primary identifier (slash-separated) — split target |
| `COL_PO`       | `"D"` | Secondary identifier — must match count with COL_PART_NO |
| `COL_WEIGHT`   | `"F"` | Weight per piece (after split, weight × pieces = total) |
| `COL_PIECES`   | `"G"` | Piece count (used for proportional redistribution) |
| `COL_VOLUME`   | `"H"` | Volume per row (L×W×H) |
| `COL_LENGTH`   | `"I"` | Box length (Constraint Convergence adjusts all three) |
| `COL_WIDTH`    | `"J"` | Box width |
| `COL_HEIGHT`   | `"K"` | Box height |
| `COL_POSTAL`   | `"L"` | Postal code (pad with leading zeros to POSTAL_LENGTH) |
| `COL_COURIER`  | `"I"` | Courier carrier column (Module 5 writes here) |
| `COL_TRACKING` | `"J"` | Tracking number (Module 5 reads here) |

## Constraint Constants

| Constant | Default | What it is | When to change |
|---|---|---|---|
| `VOLUME_MIN`      | 68.0 | Target aggregate volume lower bound | Your business needs a different range |
| `VOLUME_MAX`      | 68.1 | Target aggregate volume upper bound | Same |
| `MIN_DIMENSION`   | 0.1  | Each L/W/H must stay above this during convergence | You have a minimum-size constraint |
| `POSTAL_LENGTH`   | 5    | Pad postal codes to this many digits | Different country's postal systems |

## Runtime Populated Arrays

### `entitySuffixes` — Suffixes to strip from entity codes

Default: `Array("-30001", "-30002", "-30000", "-backup", "-temporary")`

### `routingRules` — Regex → Carrier code pairs

Default (checked in order, first match wins):
```
Array("^1Z[A-Z0-9]{16}$", "UPS")
Array("^\d{12,14}$",      "DHL")
Array("^[A-Z]{2}\d{9}$",   "DPD")
Array("^[A-Z0-9]{8,}$",    "OTHER")  ← catch-all, always last
```

## How Constraint Convergence Works (Module 4 — the core)

```
while aggregate volume not in [MIN, MAX]:
    if volume > MAX: direction = -1    # shrink all rows
    if volume < MIN: direction = +1    # grow all rows

    delta = direction × 0.01          # uniform per-row adjustment

    for each row:
        L += delta; W += delta; H += delta
        ensure L/W/H >= MIN_DIMENSION
        recompute row volume = L × W × H

    recompute aggregate volume

max iterations = 200 (guard against non-convergence)
```

**Why adjust all three dimensions uniformly?** Row proportions (shape) stay
preserved — every box shrinks/grows by the same centimeters, maintaining
its form factor.

## How Split & Redistribute Works (Module 2)

```
for each row (bottom-up, so Insert doesn't invalidate row numbers):
    if COL_PART_NO contains "/":
        split identifiers → Array parts

        insert (numParts - 1) rows below
        copy source row → each new row
        override COL_PART_NO in each copy with parts[p]
        redistribute weight/volume proportionally
        re-aggregate → absorb rounding drift on the split block
```

## How Validate & Guard Works (Module 3)

```
for each row:
    if COL_PART_NO has "/" or COL_PO has "/":
        count1 = num "/" in COL_PART_NO + 1
        count2 = num "/" in COL_PO + 1
        if count1 ≠ count2:
            red-font both columns → DO NOT SPLIT THIS ROW
```

## Second Domain Example — Server Resource Budget Convergence

| Original scenario | Second scenario |
|---|---|
| Shipping boxes in forecast table | **Servers in a data-center budget spreadsheet** |
| Aggregate volume → 68.0~68.1 CBM | **Aggregate CPU usage → 70%~75%** |
| Each row = one box with L/W/H | Each row = one server with CPU cores / MHz |
| Constraint: volume in CBM | Constraint: aggregate CPU utilization % |
| Split row = multiple 实体编号 shipments → multiple rows | **Split row = multiple tenants on one server → multiple rows** |
| Route by tracking number regex | **Route by tenant billing-code pattern** |

Same code, same constraint-convergence algorithm, different domain.

## Files

```
src/ConstraintBasedProcessor.bas   ← single source of truth
examples/sample-data.xlsx           ← fictional demo workbook
README.md                            ← Chinese-primary doc
docs/Configuration.md                ← this file
LICENSE
.gitignore
.gitattributes
```
