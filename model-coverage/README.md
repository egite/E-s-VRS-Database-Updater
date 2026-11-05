# 3D traffic model coverage, from AircraftDatabase.sqb

Counted 2026-11-03 from `AircraftDatabase.sqb` (886,242 rows). Each type was run
through the Flight Data Analysis Tool's own resolver
(`grt_data_viz/ui/adsb_model_resolver.py`, `dedicated_model_key`), so "covered"
means exactly what the 3D Flight Map would draw today.

| Rows | Aircraft |
|---|---:|
| ICAO hex only, no type or model (skipped) | 318,895 |
| With a type code or model text | 567,347 |
| Drawn by a third-party mesh | 334,554 |
| Drawn by a procedural (hand-built) mesh | 38,260 |
| No dedicated mesh (generic silhouette) | 194,533 |

Of the third-party total, 66,628 aircraft are stand-ins (drawn as a similar
type, e.g. C210 as the 182, B36T as the Bonanza).

## Files

- `type-code-counts.csv`: every ICAO type code (and code-less model string)
  with its count, the mesh it gets today and whether that is third-party,
  procedural, stand-in or generic.
- `uncovered-families-ranked.csv`: the types with no mesh, grouped into
  families one mesh can stand for, ranked by aircraft count. `stand_in_mesh`
  names an existing mesh close enough to cover the family with only a resolver
  entry; blank means it needs a new mesh. Types not yet grouped appear as
  their own one-code family.

## Procedural meshes worth replacing (aircraft drawn)

decathlon 8,094 · mooney 7,502 · bonanza 5,792 · pa46 4,293 · pc12 2,668 ·
kingair90 2,398 · huey 2,201 · vansrv 1,604 · aerocommander 823 · longez 803
