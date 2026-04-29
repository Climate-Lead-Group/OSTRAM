# A-O Parametrization Update — required to align with v17

This is a reminder of the manual updates you need to make to **A-O_Parametrization.xlsx** so that A-O matches the new v17 OSTRAM data package. None of these can be done by the merge script — they're A-O-side edits.

---

## 1. Add new Tech code: `PWRSHP` (Small Hydropower)

A-O currently has no Small Hydro code. v17 introduced PWRSHP for 5 India sub-regions only (where Small Hydro has non-zero residual capacity).

| New Tech code | Tech.Name |
|---|---|
| `PWRSHPINDEA` | Small Hydropower, India – East, region XX (investable) |
| `PWRSHPINDNE` | Small Hydropower, India – Northeast, region XX (investable) |
| `PWRSHPINDNO` | Small Hydropower, India – North, region XX (investable) |
| `PWRSHPINDSO` | Small Hydropower, India – South, region XX (investable) |
| `PWRSHPINDWE` | Small Hydropower, India – West, region XX (investable) |

Add to: A-O_Parametrization.xlsx → **Secondary Techs** sheet (and any other A-O sheets where new PWR codes need entries — check A-O_AR_Projections / A-O_AR_Model_Base_Year for reference).

A-O Tech.Name convention there is `"Small Hydropower (Power generator) India, region [EA/NE/NO/SO/WE]"` — adjust to your A-O house style.

---

## 2. Update Tech.Name for existing PWRHYD* codes (Tech codes unchanged)

A-O currently uses generic `"Hydroelectric (Power generator) [Region]"`. v17 OSTRAM uses `"Large Hydropower, [Region], region XX (investable)"`.

Tech codes are identical (`PWRHYDBGDXX`, `PWRHYDBTNXX`, `PWRHYDINDEA/NE/NO/SO/WE`, `PWRHYDLKAXX`, `PWRHYDNPLXX`) — only the Tech.Name text differs.

Decide whether you want A-O Tech.Names to track v17 OSTRAM Tech.Names verbatim or keep A-O's own house style. Either is fine — the model linkage is by Tech code, not Tech.Name.

---

## 3. Update Tech.Name for PWRCOAIND* (4 codes: INDEA/NO/SO/WE)

Tech codes unchanged. v17 collapsed Supercritical + Ultra-supercritical into a single row per parameter with merged values.

A-O Tech.Name (current convention): `"Coal (Power generator) India, region [EA/NO/SO/WE]"` — already aligned, no change needed if you're happy with this.

v17 OSTRAM uses: `"Coal-fired Power, India – [East/North/South/West], region XX (investable)"`.

If you want to align A-O Tech.Name to OSTRAM v17 verbatim, update those 4 entries.

---

## 4. Update Tech.Name for PWRNGS* (8 codes)

Same pattern as coal. Tech codes unchanged. v17 collapsed CCGT + OCGT.

Codes: `PWRNGSBGDXX`, `PWRNGSINDEA`, `PWRNGSINDNE`, `PWRNGSINDNO`, `PWRNGSINDSO`, `PWRNGSINDWE`, `PWRNGSLKAXX`, `PWRNGSMDVXX`.

A-O current: `"Natural Gas (Power generator) [Region]"` — already aligned.
v17 OSTRAM: `"Gas-fired Power, [Region], region XX (investable)"`.

Same as coal — only update if you want literal alignment.

---

## 5. Update Tech.Name for MIN India codes (4 codes)

Tech codes unchanged. v17 dedup'd 5 sub-region rows to 1 India-wide row per parameter.

| Tech code | v17 Tech.Name |
|---|---|
| `MINCOAIND` | Coal mining/import, India |
| `MINGASIND` | Natural Gas supply, India |
| `MINOILIND` | Oil supply, India |
| `MINURNIND` | Uranium supply, India |

A-O current convention: `"Coal (Mining tradable commodity) India"` etc. — already at India level, just verify the Tech.Name text matches your A-O style.

---

## What you do NOT need to do

- ❌ Add new region splits to MIN codes (A-O confirmed India-wide is correct)
- ❌ Split coal into PWRCSC/PWRCUS — we collapsed instead
- ❌ Split gas into PWRCCG/PWROCG — we collapsed instead
- ❌ Split hydro into separate Reservoir/RoR/PSP codes — we collapsed instead
- ❌ Touch any green sheets (Existing_Generation, Planned_Generation, Technology_Costs)

---

## Verification step after A-O updates

Once you've made the A-O edits, sanity-check that every Tech code in v17 Secondary_Techs has a corresponding entry in A-O Secondary Techs. Quick way:

```python
import pandas as pd
v17 = set(pd.read_excel('SOASIA_OSeMOSYS_Template_v17.xlsx', sheet_name='Secondary_Techs')['Tech'].dropna().unique())
ao  = set(pd.read_excel('A-O_Parametrization.xlsx', sheet_name='Secondary Techs')['Tech'].dropna().unique())
print('In v17 but not A-O:', v17 - ao)  # Should be empty after you add PWRSHP*
print('In A-O but not v17:', ao - v17)  # Anything here is A-O-only and that's OK
```

Expected after your edits: `In v17 but not A-O: set()` — the 5 PWRSHP* codes show up in both.
