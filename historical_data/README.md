# Historical Data

This module stores archived data and maintains change history for each sync.

Archiving
- New records are appended to daily CSV files under `historical_data/archives/YYYY-MM-DD.csv`.
- Atomic writes ensure integrity (`*.tmp` files are used and removed).

Versioning
- Each sync writes an entry to `historical_data/version_log.json` with timestamp, changed files, and a short summary.

Data Dictionary
- See `historical_data/DATA_DICTIONARY.md` for field definitions and structure.

Dependencies
- Standard Python 3 stdlib

Tests
- Run `python -m unittest discover historical_data/tests`.
