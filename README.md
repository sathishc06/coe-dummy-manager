# COE Dummy Number Manager — resolved build

## Run

```bash
python -m pip install streamlit pandas openpyxl
streamlit run coe_dummy_manager.py
```

## What this build addresses

- Aggregates source mappings across every uploaded source workbook and every worksheet.
- Detects source headers within the first 40 rows rather than assuming row 1.
- Detects target dummy columns independently for each worksheet.
- Flags any dummy number associated with multiple registration numbers and leaves its output blank.
- Flags unmatched target dummy numbers and leaves their output blank.
- Generates row-level audit records with original worksheet row numbers.
- Prevents duplicate output filenames from overwriting each other.
- Sorts whole rows and translates relative Excel formula references to their new row locations.
- Reopens output files in tests to verify the saved workbook contents.

## Test

```bash
python test_coe_app_engine.py
```

The included tests use generated Excel fixtures and check multi-source/multi-sheet mapping, non-first-row headers, conflict/unmatched handling, duplicate output names, formula-aware sorting, and output workbook reopening.

## Data privacy

When hosted, files are processed on the host server. Use only an institution-approved, access-controlled hosting environment for confidential student records. Do not upload official records to a public or unapproved service.
