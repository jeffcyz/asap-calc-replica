# Run directory layout

Git tracks the directory hierarchy through `.gitkeep` placeholders, **not the real monthly input/output files**.

```text
runs/
  AALC2201/
    2026-10/
      inputs/
        .gitkeep
        original/.gitkeep   # Original Excel workbook
        csv/.gitkeep        # Optional converted CSV data
      benchmark/.gitkeep    # SQL actual result
      output/.gitkeep       # Formula-based Excel output
      reports/.gitkeep      # Private diagnostic reports
```

The same layout exists for AALC2201, AALC2202, AALC2401, AALC2403, AALC2404, AALC2501, AALC2502, AALC2503, AALC2601.

Put period and payment_inputs under `inputs/`. The original source workbook goes in `inputs/original/`.

## Create a future run

```powershell
python scripts/setup_deals.py --all --month 2026-11
python scripts/setup_deals.py --deal AALC2201 --month 2026-11
```

## Ignore policy

The `.gitignore` deliberately allows the directory structure by retaining only `.gitkeep` markers:

```gitignore
runs/**/*
!runs/**/
!runs/**/.gitkeep
```

Any other file placed beneath `runs/` is ignored, regardless of extension. Do not use `git add -f` to override this. The repository is public; observe employer policies when using source workbooks with Copilot.
