# Data Dictionary

CSV: `historical_data/archives/YYYY-MM-DD.csv`
- `id`: Auto-increment identifier from `records` table
- `ts`: ISO 8601 timestamp when the record was produced
- `data`: JSON string with fields:
  - `line`: Production line name
  - `shift`: Shift label
  - `work_order`: Work order identifier
  - `temperature`: Temperature (°C)
  - `current`: Current (A)

Version Log: `historical_data/version_log.json`
- Array of entries:
  - `timestamp`: ISO 8601 UTC time
  - `changed_files`: Array of archive paths updated
  - `summary`: Short description of the sync
