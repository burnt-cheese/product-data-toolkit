#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""清理运行期产物（失败 CSV / 日志 / 调试快照）。

用法:
    python clean_old_runs.py            # 清 7 天前的失败行 CSV + 全部 upload.log
    python clean_old_runs.py --days 30  # 清 30 天前的
    python clean_old_runs.py --keep-log # 保留 upload.log（只清旧 CSV）
    python clean_old_runs.py --dry      # 只预览，不真删

说明:
    failed/ 下的“failed_rows_*.csv”会随每次失败上传累积，这里按修改时间清旧留新；
    upload.log / last_run.json / session.json 不在清理范围（session.json 含登录态，勿删）。
"""
import argparse
import re
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
FAILED = BASE / "failed"
LOG = BASE / "upload.log"
CSV_RE = re.compile(r"^failed_rows_\d{8}_\d{6}\.csv$")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=7, help="清理多少天前的产物（默认 7）")
    ap.add_argument("--keep-log", action="store_true", help="保留 upload.log")
    ap.add_argument("--dry", action="store_true", help="只预览不删除")
    args = ap.parse_args()

    cutoff = time.time() - args.days * 86400
    removed = []
    kept = []

    if FAILED.is_dir():
        for p in FAILED.iterdir():
            if CSV_RE.match(p.name) and p.stat().st_mtime < cutoff:
                removed.append(p)
            else:
                kept.append(p)

    if not args.keep_log and LOG.exists() and LOG.stat().st_mtime < cutoff:
        # upload.log 是累积日志；仅当其自身比 cutoff 旧才整体清（避免误删近期日志）
        removed.append(LOG)

    print(f"将清理 {args.days} 天前产物" + ("（预览模式，不删除）" if args.dry else ""))
    for p in removed:
        print(f"  DEL  {p.relative_to(BASE)}")
        if not args.dry:
            p.unlink()
    print(f"保留 {len(kept)} 个（含近期 CSV / 会话态文件），清理 {len(removed)} 个")


if __name__ == "__main__":
    main()
