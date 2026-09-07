# -*- coding: utf-8 -*-
"""SQLite run-history logging (stdlib sqlite3, no extra dependency)."""
import json
import os
import sqlite3
from datetime import datetime


SCHEMA = """
CREATE TABLE IF NOT EXISTS runs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at  TEXT NOT NULL,
    finished_at TEXT,
    input       TEXT,
    templates   TEXT,
    output      TEXT,
    report      TEXT,
    total       INTEGER,
    ok          INTEGER,
    failed      INTEGER,
    args        TEXT
);
CREATE TABLE IF NOT EXISTS files (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id     INTEGER NOT NULL REFERENCES runs(id),
    input_file TEXT NOT NULL,
    status     TEXT,
    template   TEXT,
    score      REAL,
    headings   INTEGER,
    paragraphs INTEGER,
    tables     INTEGER,
    images     INTEGER,
    warnings   TEXT,
    detail     TEXT,
    output_file TEXT,
    error      TEXT
);
CREATE INDEX IF NOT EXISTS idx_files_run ON files(run_id);
"""


def _now():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


class RunLog:
    """Appends one `runs` row plus one `files` row per input file per run.

    Writing mode also offers read-only queries (see `recent_runs` /
    `files_of_run`), so the same class is used for both recording and browsing.
    """

    def __init__(self, db_path, readonly=False):
        self.conn = None
        self.run_id = None
        self.warned = False
        if not db_path:
            return
        try:
            if readonly:
                uri = 'file:' + os.path.abspath(db_path).replace('\\', '/') + '?mode=ro'
                self.conn = sqlite3.connect(uri, uri=True)
            else:
                db_dir = os.path.dirname(os.path.abspath(db_path))
                os.makedirs(db_dir, exist_ok=True)
                self.conn = sqlite3.connect(db_path)
                self.conn.executescript(SCHEMA)
            self.conn.row_factory = sqlite3.Row
        except Exception as e:
            print(f'[警告] SQLite 历史记录不可用: {e}')
            self.conn = None

    def _safe(self, fn, *a):
        if self.conn is None:
            return
        try:
            fn(*a)
        except Exception as e:
            if not self.warned:
                print(f'[警告] 写入历史记录失败: {e}')
                self.warned = True
            try:
                self.conn.rollback()
            except Exception:
                pass

    # ------------------------------------------------------------------ #
    # writing
    # ------------------------------------------------------------------ #
    def begin(self, input_='', templates='', output='', report='', args=None):
        """Open a new run row; returns its id (or None)."""
        if self.conn is None:
            return None

        def _do():
            cur = self.conn.execute(
                'INSERT INTO runs (started_at, input, templates, output, report, args) '
                'VALUES (?,?,?,?,?,?)',
                (_now(), input_, templates, output, report,
                 json.dumps(args or [], ensure_ascii=False)))
            self.run_id = cur.lastrowid
            self.conn.commit()

        self._safe(_do)
        return self.run_id

    def log_file(self, result):
        """Record one input file's result.

        Commits per file on purpose: if the process crashes mid-batch, the
        files already processed are not lost. (Trade-off vs. one commit in
        `finish`, which is faster but loses everything on a crash.)
        """
        if self.conn is None or self.run_id is None:
            return
        st = result.stats

        def _do():
            self.conn.execute(
                'INSERT INTO files (run_id, input_file, status, template, score, '
                'headings, paragraphs, tables, images, warnings, detail, output_file, error) '
                'VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
                (self.run_id, result.input, result.status, result.template or '', result.score,
                 st.headings if st else None, st.paragraphs if st else None,
                 st.tables if st else None, st.images if st else None,
                 json.dumps(result.warnings, ensure_ascii=False),
                 result.detail or '', result.output or '', result.error or ''))
            self.conn.commit()

        self._safe(_do)

    def finish(self, total, ok, failed):
        if self.conn is None or self.run_id is None:
            return

        def _do():
            self.conn.execute(
                'UPDATE runs SET finished_at=?, total=?, ok=?, failed=? WHERE id=?',
                (_now(), total, ok, failed, self.run_id))
            self.conn.commit()

        self._safe(_do)

    # ------------------------------------------------------------------ #
    # read-only queries
    # ------------------------------------------------------------------ #
    def recent_runs(self, limit=10):
        cur = self.conn.execute('SELECT * FROM runs ORDER BY id DESC LIMIT ?', (limit,))
        return [dict(r) for r in cur.fetchall()]

    def files_of_run(self, run_id):
        cur = self.conn.execute('SELECT * FROM files WHERE run_id=? ORDER BY id', (run_id,))
        return [dict(r) for r in cur.fetchall()]

    def close(self):
        if self.conn is not None:
            try:
                self.conn.close()
            except Exception:
                pass
            self.conn = None


def _format_detail(detail):
    """Render the stored TOC-entries JSON compactly for display."""
    if not detail:
        return ''
    try:
        entries = json.loads(detail)
    except Exception:
        return ''
    shown = [f'{n} {t}' for _, n, t in entries[:8]]
    text = ' | '.join(shown)
    if len(entries) > 8:
        text += f' | …（共{len(entries)}条）'
    return text


def show_history(db_path, limit=10):
    """Print recent runs and the latest run's file details."""
    if not db_path or not os.path.exists(db_path):
        print('历史库不存在:', db_path)
        return
    log = RunLog(db_path, readonly=True)
    if log.conn is None:
        return
    try:
        runs = log.recent_runs(limit)
    except sqlite3.Error as e:
        print('读取历史库失败:', e)
        log.close()
        return
    if not runs:
        print('暂无运行记录')
        log.close()
        return

    print(f'== 最近 {len(runs)} 次运行 ==')
    for r in runs:
        print(f"  #{r['id']}  {r['started_at']}  输入={r['input']}  模板={r['templates']}  "
              f"成功/失败/总数={r['ok']}/{r['failed']}/{r['total']}  报告={r['report']}")

    latest = runs[0]['id']
    files = log.files_of_run(latest)
    print(f'== 最近一次运行（#{latest}）的文件明细 ==')
    for f in files:
        detail = _format_detail(f.get('detail'))
        print(f"  {os.path.basename(f['input_file'])} -> {f['status']}"
              f"  {f['template'] or ''}  score={f['score']:.2f}"
              f"  标题={f['headings']} 段落={f['paragraphs']} 表格={f['tables']} 图片={f['images']}"
              f"  {f['warnings'] or ''}"
              + (f"  目录={detail}" if detail else ''))
    log.close()
