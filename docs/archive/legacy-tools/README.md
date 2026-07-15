# Legacy tools

This directory preserves obsolete maintenance and analysis scripts for provenance. The
files are historical code, not supported commands, and must not be executed against the
current model tree.

| Archived file | Former path | Reason retired |
|---|---|---|
| `concat_all_scenarios_merge.py` | `t1_confection/concat_all_scenarios.py` | Its merge strategy multiplied rows at mixed parameter granularities; `t1_confection/concat_all_scenarios_2.py` is the maintained stack-based successor. |
| `Z_AUX_united_regions.py` | `t1_confection/Z_AUX_united_regions.py` | It is a hard-coded Brazil-region transformation with no inbound production reference in the current South/Southeast Asia workflow. |
