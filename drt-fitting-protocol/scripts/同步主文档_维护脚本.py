# -*- coding: utf-8 -*-
"""把包内协议正文反向同步到**单文件精简版主文档**（`<doc-dir>/DRT拟合原则_精简版.md`）。

背景（v1.12）：协议正文已精简（去修订记录与版本标注），因此
  * 受维护的单文件文档 = `DRT拟合原则_精简版.md`（= 正文 + 附录 A/B/D，无版本史）；
  * `DRT拟合原则.md`（全量版，含附录 C 版本变更与附录 B.2 修订证据）**是历史归档，本脚本不再维护它** ——
    不要对它跑 `--apply`，否则会用精简正文覆盖归档。

变换（**全部可逆**，便于逐字节核对）：
  S1 `@references/附录B_实证依据.md`   -> `附录 B`
  S2 `@references/附录D_归档工具箱备份.md` -> `附录 D`
  S3 `scripts/DRT归档工具箱.py`        -> `DRT归档工具箱.py`（主文档与工具箱同目录）
另插入主文档特有的"精简版自述 + 资源位置"块；核对时按正向合成比对（不做逆变换：
S3 有歧义 —— `DRT归档工具箱.py` 在正文中本来就存在）。

用法（默认 dry-run）：
  python scripts/同步主文档_维护脚本.py --doc-dir "<主文档目录>"            # 干跑
  python scripts/同步主文档_维护脚本.py --doc-dir "<主文档目录>" --apply    # 备份 + 写盘
"""
import os
import re
import sys
import hashlib
import argparse
import shutil
import datetime

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))      # 技能包根
SKILL = os.path.join(BASE, "SKILL.md")
APPX_B = os.path.join(BASE, "references", "附录B_实证依据.md")
APPX_D = os.path.join(BASE, "references", "附录D_归档工具箱备份.md")
TOOL = os.path.join(BASE, "scripts", "DRT归档工具箱.py")

DOC_NAME = "DRT拟合原则_精简版.md"
TOOL_NAME = "DRT归档工具箱.py"

NOTE = """> **本文件是精简版**：不含修订记录与版本标注（完整版见同目录 `DRT拟合原则.md`）；**条款与硬门槛 H1–H21（含 H1′/H5′/H7′/H8′）一字未减**。
> **资源位置（单文件自包含版）**：归档工具箱 = 与本文档同目录的 `DRT归档工具箱.py`（与技能包 `scripts/DRT归档工具箱.py` 逐字节相同）；**附录 B**（实证依据）、**附录 D**（工具箱全文）都在本文档内；`scripts/同步附录D_维护脚本.py`、`scripts/自检_协议包.py`、`scripts/自检_工具箱回归.py`、`scripts/同步主文档_维护脚本.py` 属技能包 `scripts/`（依赖包的目录结构，不在本文档目录）。
"""


def sha(p):
    with open(p, "rb") as fh:
        return hashlib.sha256(fh.read()).hexdigest()


def read(p):
    return open(p, encoding="utf-8").read()


def clean_tail(s):
    """去掉片段自带的页脚与尾部 `---`（附录 B/D 文件末尾都带，若不清掉会堆出重复分隔线与页脚）。"""
    s = re.sub(r"\n*---\s*\n\*DRT 拟合原则[\s\S]*$", "", s)
    s = re.sub(r"(?:\s*---\s*)+$", "", s)
    return s.rstrip()


def compose_body(skill_text, ver):
    """SKILL.md -> 主文档正文段（§0–§8 + 附录 A）。先替换、后插 NOTE（顺序不能反）。"""
    body = skill_text[skill_text.index("# DRT 拟合原则"):].rstrip() + "\n"
    body = body.replace("# DRT 拟合原则 %s" % ver, "# DRT 拟合原则 %s（精简版）" % ver, 1)
    subs = [("`@references/附录B_实证依据.md`", "**附录 B**"),
            ("`@references/附录D_归档工具箱备份.md`", "**附录 D**"),
            ("`scripts/DRT归档工具箱.py`", "`DRT归档工具箱.py`")]
    counts = []
    for a, b in subs:
        counts.append((a, body.count(a)))
        body = body.replace(a, b)
    pos = body.index("\n---\n")
    body = body[:pos + 1] + NOTE + body[pos + 1:]
    return clean_tail(body), counts


def build(doc_dir):
    ver = re.search(r'^PROTOCOL_VERSION = "(v[\d.]+)"', read(TOOL), re.M).group(1)
    body, counts = compose_body(read(SKILL), ver)
    app_b, app_d = clean_tail(read(APPX_B)), clean_tail(read(APPX_D))
    foot = ("\n\n---\n\n*DRT 拟合原则 %s（精简版） — 建议在每次分析开始时把本文件全文连同数据一并提供给 AI。*\n"
            "*文件名不带版本号，本文档始终为最新版。*\n" % ver)
    doc = body + "\n\n---\n\n" + app_b + "\n\n---\n\n" + app_d + foot
    return doc, dict(counts=counts, ver=ver)


def verify(doc, doc_p):
    """正向可复现 + 附录 D 逐字节 + 章节顺序 + 编码 + 修订痕迹。"""
    ver = re.search(r'^PROTOCOL_VERSION = "(v[\d.]+)"', read(TOOL), re.M).group(1)
    expected, _ = compose_body(read(SKILL), ver)
    actual = doc[:doc.index("\n\n---\n\n## 附录 B")]
    print("正向可复现  : 文档正文段 == compose_body(SKILL.md) ->",
          "PASS" if actual.rstrip() == expected.rstrip() else "FAIL")
    d = doc[doc.index("## 附录 D · 归档工具箱"):]
    code = d[d.index("```python\n") + len("```python\n"):d.rindex("\n```")]
    tool = read(TOOL)
    print("附录 D 正文 : 与工具箱逐字节相同 ->", "PASS" if code.rstrip("\n") == tool.rstrip("\n") else "FAIL")
    decl = re.search(r"`sha256 = ([0-9a-f]{64})`", d).group(1)
    print("附录 D 声明 : sha256 == 工具箱实际 ->", "PASS" if decl == sha(TOOL) else "FAIL", "(%s…)" % decl[:12])
    heads = []
    for h in re.findall(r"^## (\d+|附录 [A-D])", doc, re.M):
        if h not in heads:
            heads.append(h)
    print("章节顺序    :", " → ".join(heads))
    nofence = re.sub(r"```[\s\S]*?```", "", doc)
    bad = {p: len(re.findall(p, nofence)) for p in (r"## 0\.2 ", r"## 0\.3 ", r"§0\.2", r"§0\.3", r"（v1\.\d+")}
    bad = {k: v for k, v in bad.items() if v}
    print("修订痕迹    :", "PASS（0 命中）" if not bad else "FAIL " + str(bad))
    if os.path.exists(doc_p):
        b = open(doc_p, "rb").read()
        print("编码/换行   : BOM=%s CRLF=%d" % (b[:3] == b"\xef\xbb\xbf", b.count(b"\r\n")))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--doc-dir", required=True, help="主文档所在目录")
    ap.add_argument("--apply", action="store_true", help="备份并写盘（默认只干跑）")
    ap.add_argument("--backup-tag", default=None, help="备份目录后缀，默认 v<旧版本>_<日期>")
    a = ap.parse_args()
    doc_dir = os.path.abspath(a.doc_dir)
    doc_p = os.path.join(doc_dir, DOC_NAME)
    tool_p = os.path.join(doc_dir, TOOL_NAME)
    if not os.path.isdir(doc_dir):
        sys.exit("ERROR: 目录不存在 " + doc_dir)
    doc, info = build(doc_dir)
    old = read(doc_p) if os.path.exists(doc_p) else ""
    old_ver = (re.search(r"^# DRT 拟合原则 (v[\d.]+)", old, re.M) or [None, "无"])[1] if old else "无"
    print("目标文档    :", doc_p, "（现 %s）" % (("%d 字节，%s" % (len(old.encode()), old_ver)) if old else "不存在"))
    print("工具箱      :", tool_p, "（现 %s，%d 字节）" % (sha(tool_p)[:12], os.path.getsize(tool_p)))
    print("目标状态    : 文档 %d 字节 / 工具箱 %d 字节（包内 %s）"
          % (len(doc.encode()), os.path.getsize(TOOL), info["ver"]))
    print("替换计数    :", ", ".join("%s×%d" % (x, n) for x, n in info["counts"]))
    same_doc, same_tool = (doc == old), os.path.exists(tool_p) and sha(tool_p) == sha(TOOL)
    if same_doc and same_tool:
        print("\n已是目标状态（文档与工具箱都与包内一致）→ 不写盘、不备份。")
        verify(doc, doc_p)
        return 0
    if not a.apply:
        print("\n[DRY-RUN] 未写盘。加 --apply 执行（会先备份旧文档/旧工具箱到 _backup_<tag>/）。")
        verify(doc, doc_p)
        return 0
    tag = a.backup_tag or ("%s_%s" % (old_ver, datetime.date.today().isoformat()))
    bdir = os.path.join(doc_dir, "_backup_" + tag)
    os.makedirs(bdir, exist_ok=True)
    lines = []
    for fn in sorted(os.listdir(doc_dir)):
        p = os.path.join(doc_dir, fn)
        if not os.path.isfile(p) or fn.startswith("_") or not fn.endswith((".md", ".py")):
            continue
        shutil.copy2(p, os.path.join(bdir, fn))
        lines.append("%s  %s  %d" % (sha(p), fn, os.path.getsize(p)))
    open(os.path.join(bdir, "_manifest_sha256.txt"), "w", encoding="utf-8", newline="\n").write(
        "sha256  name  bytes\n" + "\n".join(lines) + "\n")
    print("\n已备份 %d 个文件 -> %s" % (len(lines), bdir))
    open(doc_p, "w", encoding="utf-8", newline="").write(doc)
    shutil.copy2(TOOL, tool_p)
    print("已写入 %s（%d 字节）与 %s（%d 字节）" % (DOC_NAME, len(doc.encode()), TOOL_NAME,
                                                  os.path.getsize(tool_p)))
    verify(doc, doc_p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
