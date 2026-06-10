#!/usr/bin/env python3
"""
FOLDER ORGANISER — dedupe, archive, auto-index. SAFE BY DEFAULT (dry-run).
Usage:
  python folder_organiser.py <folder>            # dry-run report only
  python folder_organiser.py <folder> --apply    # archive backups (move, never delete)
  python folder_organiser.py <folder> --index    # write FOLDER_INDEX.md

Functions:
  1. DEDUPE  — SHA256 hash of every file; identical content reported in
               Duplicates_Report.md (originals NEVER deleted automatically).
  2. ARCHIVE — files matching backup patterns (*.bak, *.bak2, *.audit_bak,
               *.corrupt_*, *.pre*, *~, *.old) moved to _Archive/ with --apply.
  3. INDEX   — FOLDER_INDEX.md tree of all folders + files with sizes.
"""
import hashlib, os, shutil, sys
from collections import defaultdict

BACKUP_SUFFIX = (".bak", ".bak2", ".old", "~")
BACKUP_CONTAINS = (".audit_bak", ".corrupt_", ".pre")

def is_backup(fn):
    low = fn.lower()
    return low.endswith(BACKUP_SUFFIX) or any(p in low for p in BACKUP_CONTAINS)

def sha256(path, block=1 << 20):
    h = hashlib.sha256()
    try:
        with open(path, "rb") as f:
            while True:
                b = f.read(block)
                if not b: break
                h.update(b)
        return h.hexdigest()
    except OSError:
        return None

def main(root, apply=False, index=False):
    root = os.path.abspath(root)
    hashes = defaultdict(list); backups = []; total = 0
    for dirpath, dirs, fns in os.walk(root):
        dirs[:] = [d for d in dirs if not d.startswith((".", "_Archive"))]
        for fn in fns:
            p = os.path.join(dirpath, fn); total += 1
            if is_backup(fn): backups.append(p)
            h = sha256(p)
            if h: hashes[(h, os.path.getsize(p))].append(p)

    dupes = {k: v for k, v in hashes.items() if len(v) > 1}
    rep = [f"# DUPLICATES REPORT — {root}", f"Files scanned: {total}",
           f"Duplicate groups (identical SHA256): {len(dupes)}", ""]
    wasted = 0
    for (h, size), paths in sorted(dupes.items(), key=lambda x: -x[0][1]):
        wasted += size * (len(paths) - 1)
        rep.append(f"## {size//1024} KB × {len(paths)} copies  (sha {h[:12]}…)")
        rep += [f"  - {os.path.relpath(p, root)}" for p in sorted(paths)]
        rep.append("")
    rep.insert(3, f"Reclaimable space if deduped: {wasted // (1<<20)} MB (NO files were deleted — review first)")
    open(os.path.join(root, "Duplicates_Report.md"), "w").write("\n".join(rep))
    print(f"✓ Duplicates_Report.md — {len(dupes)} duplicate groups, {wasted//(1<<20)} MB reclaimable")

    if backups:
        arc = os.path.join(root, "_Archive", "auto_backups")
        print(f"{'Moving' if apply else 'DRY-RUN would move'} {len(backups)} backup files → _Archive/auto_backups/")
        if apply:
            os.makedirs(arc, exist_ok=True)
            for p in backups:
                dest = os.path.join(arc, os.path.basename(p))
                i = 1
                while os.path.exists(dest):
                    dest = os.path.join(arc, f"{os.path.basename(p)}.{i}"); i += 1
                shutil.move(p, dest)
    else:
        print("No backup-pattern files found.")

    if index:
        lines = [f"# FOLDER INDEX — {os.path.basename(root)}", ""]
        cnt = 0
        for dirpath, dirs, fns in os.walk(root):
            dirs[:] = sorted(d for d in dirs if not d.startswith("."))
            rel = os.path.relpath(dirpath, root)
            depth = 0 if rel == "." else rel.count(os.sep) + 1
            if rel != ".":
                lines.append(f"{'  '*depth}📁 {os.path.basename(dirpath)}/ ({len(fns)} files)")
            for fn in sorted(fns):
                cnt += 1
                if depth <= 1:
                    sz = os.path.getsize(os.path.join(dirpath, fn)) // 1024
                    lines.append(f"{'  '*(depth+1)}- {fn} ({sz} KB)")
        lines.append(f"\nTOTAL FILES: {cnt}")
        open(os.path.join(root, "FOLDER_INDEX.md"), "w").write("\n".join(lines))
        print(f"✓ FOLDER_INDEX.md — {cnt} files indexed")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    main(sys.argv[1], apply="--apply" in sys.argv, index="--index" in sys.argv)
