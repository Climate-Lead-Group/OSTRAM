# TRN ResidualCapacity Fix — Diff Log

- Mode: **min** (`TotalAnnualMinCapacityInvestment` if min, `TotalAnnualMaxCapacity` if max)
- Cutoff year: **2023**
- Techs split (residual flattened, deltas → `TotalAnnualMinCapacityInvestment`): **11**
- Techs with magnitude-only correction (already flat in input, level adjusted from reference): **1**
- Techs skipped (no change needed): **6**
- Cell changes from magnitude correction: **28**
- Cell changes from residual splitting: **281**

## Skipped (no change needed — already flat & magnitude OK)
- `TRNBGDXXINDNE`
- `TRNINDEAINDNE`
- `TRNINDEAINDNO`
- `TRNINDEAINDSO`
- `TRNINDNEINDNO`
- `TRNMDVXXINDSO`

## Magnitude-only corrections (input was already flat; level changed from reference)
- `TRNBGDXXINDEA`: 1.000 → **2.500** (flat across all years)

## Per-tech commissioning schedules (residual splits)

### `TRNBTNXXBGDXX`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **0.000** _(base from input)_
- Post-2023 commissionings (moved to `TotalAnnualMinCapacityInvestment`):
  - 2035: +0.750

### `TRNBTNXXINDEA`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **5.526** _(base from reference)_
- Commissioning deltas derived from reference profile
- (no post-cutoff commissionings derived)

### `TRNBTNXXINDNE`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **0.110** _(base from reference)_
- Commissioning deltas derived from reference profile
- (no post-cutoff commissionings derived)

### `TRNINDEAINDWE`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **21.190** _(base from input)_
- Post-2023 commissionings (moved to `TotalAnnualMinCapacityInvestment`):
  - 2025: +1.600

### `TRNINDEANPLXX`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **1.300** _(base from reference)_
- Commissioning deltas derived from reference profile
- (no post-cutoff commissionings derived)

### `TRNINDNOINDWE`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **36.720** _(base from input)_
- Post-2023 commissionings (moved to `TotalAnnualMinCapacityInvestment`):
  - 2024: +1.600
  - 2026: +8.400
  - 2028: +4.200

### `TRNINDNONPLXX`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **1.200** _(base from reference)_
- Commissioning deltas derived from reference profile
- (no post-cutoff commissionings derived)

### `TRNINDSOINDWE`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **18.120** _(base from input)_
- Post-2023 commissionings (moved to `TotalAnnualMinCapacityInvestment`):
  - 2024: +4.200
  - 2025: +4.200
  - 2026: +1.600

### `TRNINDSOLKAXX`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **0.000** _(base from input)_
- Post-2023 commissionings (moved to `TotalAnnualMinCapacityInvestment`):
  - 2030: +0.500

### `TRNLKAXXMDVXX`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **0.000** _(base from input)_
- Post-2023 commissionings (moved to `TotalAnnualMinCapacityInvestment`):
  - 2030: +0.500

### `TRNNPLXXBGDXX`
- Pre-2023 stock (kept in `ResidualCapacity`, flat across all years): **0.040** _(base from reference)_
- Commissioning deltas derived from reference profile
- Post-2023 commissionings (moved to `TotalAnnualMinCapacityInvestment`):
  - 2033: +1.000
