# Real-Time Monitoring

This module handles live acquisition and visualization during runtime.

Usage
- Launch the GUI: `python 3rows.py`
- Select serial port and sensor addresses; click `開始` to start recording.

Lifecycle
- Temporary files reside under `real_time_monitoring/temp/`.
- While recording, a `recording.lock` file is touched periodically.
- On stop/reset (all lines stopped) and on application exit, all files in `real_time_monitoring/temp/` are automatically deleted.

Contents
- `temp/` runtime artifacts: `ems.db`, `recording.lock` and transient files.
- The GUI writes per-tick records into `ems.db` for downstream sync.

Dependencies
- PyQt5
- pyserial (optional for serial port)
- modbus-tk (optional for Modbus RTU)

Tests
- See `historical_data/tests` for sync pipeline tests (realtime cleanup is verified by integration).
