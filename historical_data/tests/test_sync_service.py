import os
import time
import json
from pathlib import Path
import unittest

from sync_service import SyncService

class TestSyncService(unittest.TestCase):
    def setUp(self):
        self.svc = SyncService(interval=1)
        self.temp = self.svc.var_dir
        self.archives = self.svc.exports_dir
        self.version_log = self.svc.version_log
        self.temp.mkdir(parents=True, exist_ok=True)
        self.archives.mkdir(parents=True, exist_ok=True)
        if self.version_log.exists():
            self.version_log.unlink()

    def tearDown(self):
        for d in [self.archives, self.temp]:
            if d.exists():
                for p in d.iterdir():
                    try:
                        p.unlink()
                    except Exception:
                        pass

    def test_is_recording_flag(self):
        if self.svc.recording_flag.exists():
            self.svc.recording_flag.unlink()
        self.assertFalse(self.svc._is_recording())
        self.svc.recording_flag.parent.mkdir(parents=True, exist_ok=True)
        self.svc.recording_flag.write_text('1', encoding='utf-8')
        self.assertTrue(self.svc._is_recording())
        os.utime(self.svc.recording_flag, (time.time()-120, time.time()-120))
        self.assertFalse(self.svc._is_recording())

    def test_write_exports_and_version_log(self):
        now = time.time()
        recs = [
            {'id': 1, 'ts': '2025-01-01T00:00:00+00:00', 'data': '{"a":1}'},
            {'id': 2, 'ts': '2025-01-01T00:00:01+00:00', 'data': '{"a":2}'}
        ]
        paths = self.svc._write_exports(recs)
        self.assertTrue(paths)
        self.svc._commit_and_push(paths)
        self.assertTrue(self.version_log.exists())
        log = json.loads(self.version_log.read_text(encoding='utf-8'))
        self.assertTrue(len(log) >= 1)

if __name__ == '__main__':
    unittest.main()
