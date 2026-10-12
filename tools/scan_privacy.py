#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""发布前隐私体检：通用规则 + 运行时禁词。

设计要点：本脚本会进公开仓库，因此**不得**把任何具体数据标识写进源码 ——
那等于把要隐藏的东西公开。具体标识一律通过环境变量传入：

    PRIVACY_DENY   逗号分隔的额外禁词（大小写不敏感），例如样品名、批次名、本机路径片段。
                   CI 里用仓库变量传入：Settings -> Secrets and variables -> Actions -> Variables。

检查项（对 zip 或目录均可）：
    R1 数据类扩展名（.txt/.csv/.xlsx/...）—— 包内不应出现任何实测数据文件
    R2 绝对路径（盘符式 C:\\ / D:/，排除 URL 里的 "s:/"）
    R3 家目录（C:\\Users\\、/home/、/Users/）
    R4 邮箱地址
    R5 中国大陆手机号
    R6 NUL 字节
    R7 运行时禁词 PRIVACY_DENY（含 zip 原始字节，覆盖中央目录里的文件名表）

用法：
    python tools/scan_privacy.py dist/drt-fitting-protocol.zip [更多目标...]
退出码：0 通过；1 有命中。
"""
import os
import re
import sys
import zipfile

DATA_EXT = (".txt", ".csv", ".xlsx", ".xls", ".dat", ".z", ".mpr", ".nox", ".mpt", ".bin")
RULES = [
    ("R2 绝对路径", re.compile(r"(?<![A-Za-z0-9])[A-Za-z]:[\\/]")),
    ("R3 家目录", re.compile(r"C:\\Users\\|/home/|/Users/")),
    ("R4 邮箱", re.compile(r"[\w.+-]+@[\w-]+\.[A-Za-z]{2,}")),
    ("R5 手机号", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")),
]


def deny_tokens():
    raw = os.environ.get("PRIVACY_DENY", "")
    return [t.strip() for t in re.split(r"[,;]", raw) if t.strip()]


def iter_files(target):
    """产出 (显示名, 字节内容)。"""
    if target.lower().endswith(".zip"):
        with zipfile.ZipFile(target) as z:
            for n in z.namelist():
                yield n, z.read(n)
    else:
        for dp, dn, fn in os.walk(target):
            if os.path.basename(dp) in ("__pycache__", ".git"):
                continue
            for f in sorted(fn):
                fp = os.path.join(dp, f)
                yield os.path.relpath(fp, target).replace("\\", "/"), open(fp, "rb").read()


def scan(target, deny):
    hits = []
    for name, data in iter_files(target):
        ext = os.path.splitext(name)[1].lower()
        if ext in DATA_EXT:
            hits.append((name, 0, "R1 数据类扩展名", name))
        if b"\x00" in data:
            hits.append((name, 0, "R6 NUL 字节", "%d 个" % data.count(b"\x00")))
        text = data.decode("utf-8", "ignore")
        for ln, line in enumerate(text.split("\n"), 1):
            for tag, rx in RULES:
                m = rx.search(line)
                if m:
                    hits.append((name, ln, tag, m.group(0)))
            low = line.lower()
            for t in deny:
                if t.lower() in low:
                    hits.append((name, ln, "R7 禁词", t))
    # zip 原始字节（中央目录/文件名表不在逐条文本里）
    if target.lower().endswith(".zip"):
        raw = open(target, "rb").read()
        low = raw.lower()
        for t in deny:
            if t.lower().encode("utf-8") in low:
                hits.append(("<zip 原始字节>", 0, "R7 禁词", t))
    return hits


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    deny = deny_tokens()
    print("隐私体检：禁词 %d 个（来自 PRIVACY_DENY）" % len(deny))
    total = 0
    for target in argv:
        hits = scan(target, deny)
        total += len(hits)
        print("  %s -> %s（命中 %d）" % (target, "PASS" if not hits else "FAIL", len(hits)))
        for name, ln, tag, frag in hits[:40]:
            print("      [%s] %s:%s  %s" % (tag, name, ln or "-", frag[:80]))
    print("隐私体检结论：%s（命中 %d）" % ("PASS" if total == 0 else "FAIL", total))
    return 0 if total == 0 else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
