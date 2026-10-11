# -*- coding: utf-8 -*-
"""把 references/附录D_归档工具箱备份.md 的代码段与 scripts/DRT归档工具箱.py 重新对齐（逐字节相同），
并同步更新头部声明的 sha256 / 字节数 / 行数 / 版本号。

版本号**从工具箱读取**（PROTOCOL_VERSION），不在本脚本里硬编码 —— 否则协议升版时附录 D 的页脚会
停在旧版本（§8.0.6 B3 的同类事故：三份交付物分别写着 v1.4/v1.4/v1.5）。
"""
import os
import re
import sys
import hashlib

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 机器无关：本脚本所在技能包的根
TOOL = os.path.join(BASE, "scripts", "DRT归档工具箱.py")
BACKUP = os.path.join(BASE, "references", "附录D_归档工具箱备份.md")
if not os.path.exists(TOOL) or not os.path.exists(BACKUP):
    sys.exit("ERROR: 找不到工具箱或附录 D —— 请把本脚本放在技能包的 scripts/ 目录下运行")

code = open(TOOL, encoding="utf-8").read()
m = re.search(r'^PROTOCOL_VERSION = "(v[\d.]+)"', code, re.M)
if not m:
    sys.exit("ERROR: 工具箱里找不到 PROTOCOL_VERSION，无法同步附录 D 的版本串")
VER = m.group(1)
sha = hashlib.sha256(code.encode("utf-8")).hexdigest()
nb = os.path.getsize(TOOL)
nl = code.count("\n") + (0 if code.endswith("\n") else 1)
text = open(BACKUP, encoding="utf-8").read()

i = text.index("```python\n") + len("```python\n")
j = text.rindex("\n```")
head, tail = text[:i], text[j:]

# 头部：正文说明 + 声明的 hash/字节/行数 + 版本串（版本号取自工具箱唯一常量）
head = re.sub(r"`sha256 = [0-9a-f]{64}`", "`sha256 = %s`" % sha, head)
head = re.sub(r"（\d+ 字节，\d+ 行）", "（%d 字节，%d 行）" % (nb, nl), head)
head = re.sub(r"DRT 拟合原则 v[\d.]+", "DRT 拟合原则 %s" % VER, head)
tail = re.sub(r"DRT 拟合原则 v[\d.]+", "DRT 拟合原则 %s" % VER, tail)
tail = re.sub(r"`summary`（§8.4 汇总表）", "`summary`（§8.4 汇总表 + §8.3 DRT_batch_summary.csv）", tail)

new = head + code.rstrip("\n") + "\n" + tail
open(BACKUP, "w", encoding="utf-8", newline="").write(new)

# 复核：抽回代码段逐字节比对 + 版本串一致
t2 = open(BACKUP, encoding="utf-8").read()
a = t2.index("```python\n") + len("```python\n")
b = t2.rindex("\n```")
extracted = t2[a:b]
print("工具箱版本常量    :", VER)
print("工具 sha256      :", sha)
print("工具 字节/行数    :", nb, nl)
print("附录代码段逐字节相同:", extracted.rstrip("\n") == code.rstrip("\n"),
      "| 长度", len(extracted), "vs", len(code))
print("附录声明的 sha256 :", re.search(r"`sha256 = ([0-9a-f]{64})`", t2).group(1))
print("附录声明的字节行数:", re.search(r"（\d+ 字节，\d+ 行）", t2).group(0))
for lbl, seg in (("头部", t2[:a]), ("页脚", t2[b:])):
    vs = sorted(set(re.findall(r"DRT 拟合原则 (v[\d.]+)", seg)))
    print("附录%s版本串     : %s %s" % (lbl, vs, "OK" if vs in ([VER], []) else "!!! 与工具箱不一致"))
print("附录文件大小      :", os.path.getsize(BACKUP))
