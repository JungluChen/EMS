import os
import time
import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime, timezone
import subprocess
import sqlite3
import argparse

class SyncService:
    def __init__(self, interval=5, db_path=None):
        self.root = Path(__file__).parent
        self.exports_dir = self.root / 'historical_data' / 'archives'
        self.logs_dir = self.root / 'historical_data' / 'logs'
        self.var_dir = self.root / 'real_time_monitoring' / 'temp'
        self.state_path = self.var_dir / 'sync_state.json'
        self.recording_flag = self.var_dir / 'recording.lock'
        self.version_log = self.root / 'historical_data' / 'version_log.json'
        self.interval = int(interval)
        self.db_path = db_path or os.environ.get('DB_PATH') or str(self.var_dir / 'ems.db')
        self.git_owner = os.environ.get('GIT_OWNER') or 'JungluChen'
        self.git_repo = os.environ.get('GIT_REPO') or 'EMS'
        self.git_branch = os.environ.get('GIT_BRANCH') or 'main'
        self._prepare_dirs()
        self._setup_logging()

    def _prepare_dirs(self):
        for d in [self.exports_dir, self.logs_dir, self.var_dir]:
            d.mkdir(parents=True, exist_ok=True)

    def _setup_logging(self):
        log_file = self.logs_dir / 'sync.log'
        handler = RotatingFileHandler(str(log_file), maxBytes=512000, backupCount=3)
        fmt = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
        handler.setFormatter(fmt)
        logger = logging.getLogger('sync')
        logger.setLevel(logging.INFO)
        logger.addHandler(handler)
        self.log = logger

    def _load_last_sync(self):
        if self.state_path.exists():
            try:
                with open(self.state_path, 'r', encoding='utf-8') as f:
                    obj = json.load(f)
                    return obj.get('last_sync')
            except Exception:
                return None
        return None

    def _save_last_sync(self, ts_iso):
        try:
            with open(self.state_path, 'w', encoding='utf-8') as f:
                json.dump({'last_sync': ts_iso}, f)
        except Exception as e:
            self.log.error(f'save_last_sync failed: {e}')

    def _connect_db(self):
        if not self.db_path or not Path(self.db_path).exists():
            return None
        try:
            conn = sqlite3.connect(self.db_path)
            return conn
        except Exception as e:
            self.log.error(f'db connect failed: {e}')
            return None

    def _extract_new_records(self, since_iso):
        conn = self._connect_db()
        if not conn:
            return []
        try:
            cur = conn.cursor()
            if since_iso:
                cur.execute('SELECT id, ts, data FROM records WHERE ts > ? ORDER BY ts ASC', (since_iso,))
            else:
                cur.execute('SELECT id, ts, data FROM records ORDER BY ts ASC')
            rows = cur.fetchall()
            conn.close()
            recs = []
            for rid, ts, data in rows:
                recs.append({'id': rid, 'ts': ts, 'data': data})
            return recs
        except Exception as e:
            self.log.error(f'db query failed: {e}')
            try:
                conn.close()
            except Exception:
                pass
            return []

    def _write_exports(self, records):
        if not records:
            return []
        changed = set()
        for r in records:
            try:
                dt = datetime.fromisoformat(r['ts'])
            except Exception:
                dt = datetime.now(timezone.utc)
            fname = dt.strftime('%Y-%m-%d') + '.csv'
            out_path = self.exports_dir / fname
            tmp_path = out_path.with_suffix('.csv.tmp')
            header = not out_path.exists()
            with open(tmp_path, 'a', newline='') as f:
                if header:
                    f.write('id,ts,data\n')
                f.write(f"{r['id']},{r['ts']},{r['data']}\n")
            if out_path.exists():
                with open(tmp_path, 'rb') as tf, open(out_path, 'ab') as of:
                    of.write(tf.read())
                tmp_path.unlink(missing_ok=True)
            else:
                tmp_path.rename(out_path)
            changed.add(str(out_path))
        return list(changed)

    def _git(self, args):
        try:
            subprocess.run(['git'] + args, cwd=str(self.root), check=True, capture_output=True)
            return True
        except subprocess.CalledProcessError as e:
            msg = e.stderr.decode('utf-8', errors='ignore') if e.stderr else str(e)
            self.log.warning(f'git {" ".join(args)} failed: {msg}')
            return False

    def _commit_and_push(self, paths):
        if not paths:
            return
        ok_add = self._git(['add'] + paths)
        if not ok_add:
            return
        msg = f"sync: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ok_commit = self._git(['commit', '-m', msg])
        if not ok_commit:
            return
        self._git(['push', 'origin', self.git_branch])
        try:
            self._update_version_log(paths)
        except Exception as e:
            self.log.warning(f'version log update failed: {e}')

    def _update_version_log(self, paths):
        entry = {
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'changed_files': paths,
            'summary': f"synced {len(paths)} file(s)"
        }
        log = []
        if self.version_log.exists():
            try:
                with open(self.version_log, 'r', encoding='utf-8') as f:
                    log = json.load(f)
            except Exception:
                log = []
        log.append(entry)
        self.version_log.parent.mkdir(parents=True, exist_ok=True)
        with open(self.version_log, 'w', encoding='utf-8') as f:
            json.dump(log, f, ensure_ascii=False, indent=2)

    def run_once(self):
        since = self._load_last_sync()
        recs = self._extract_new_records(since)
        if not recs:
            self.log.info('no new records')
            return
        paths = self._write_exports(recs)
        self._commit_and_push(paths)
        now_iso = datetime.now(timezone.utc).isoformat()
        self._save_last_sync(now_iso)
        self.log.info(f'synced {len(recs)} records')

    def run_forever(self):
        while True:
            try:
                if self._is_recording():
                    self.run_once()
                else:
                    self.log.info('idle: recording not active, skip')
            except Exception as e:
                self.log.error(f'run_once error: {e}')
            time.sleep(self.interval)

    def _is_recording(self):
        try:
            if not self.recording_flag.exists():
                return False
            st = self.recording_flag.stat()
            age = time.time() - float(st.st_mtime)
            threshold = max(self.interval * 2, 15)
            return age < threshold
        except Exception:
            return False

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--once', action='store_true')
    parser.add_argument('--interval', type=int, default=5)
    parser.add_argument('--db-path', type=str, default=None)
    args = parser.parse_args()
    svc = SyncService(interval=args.interval, db_path=args.db_path)
    if args.once:
        svc.run_once()
    else:
        svc.run_forever()

if __name__ == '__main__':
    main()
