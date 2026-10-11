# -*- coding: utf-8 -*-
"""协议包自检（v1.10 新增）—— 改动协议正文 / 工具箱 / 附录后跑一次。

检查项：
  P1 文档表格结构：SKILL.md、附录 B 的每个 Markdown 表格"每行列数 == 表头列数"（§8.0.6 F4）；
  P2 版本串一致：SKILL.md 标题 / front-matter version / 工具箱 PROTOCOL、PROTOCOL_VERSION / 附录 D 页脚；
  P3 附录 D 与工具箱**逐字节相同**，且头部声明的 sha256 / 字节数 / 行数与实际一致；
  P4 §0.2 的 `common_window()` 参考实现行为正确（对频率取 max，不是最大索引）；
  P5 工具箱五个子命令可调用、`--help` 正常。

用法：python scripts/自检_协议包.py
退出码：0 = 全部 PASS；1 = 有 FAIL。
"""
import os
import re
import sys
import json
import hashlib
import subprocess

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL = os.path.join(BASE, "SKILL.md")
APX_B = os.path.join(BASE, "references", "附录B_实证依据.md")
APX_D = os.path.join(BASE, "references", "附录D_归档工具箱备份.md")
TOOL = os.path.join(BASE, "scripts", "DRT归档工具箱.py")
ok = True


def say(tag, cond, msg):
    global ok
    ok = ok and bool(cond)
    print("  [%s] %s" % ("PASS" if cond else "FAIL", msg))


# ---------------- P1 表格结构 ----------------
def split_row(line):
    s = line.replace(r"\|", "\x00").strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


SEP = re.compile(r"^\s*\|[\s:|-]+\|\s*$")


def check_tables(path):
    lines = open(path, encoding="utf-8-sig").read().split("\n")
    fence = False
    bad, nt, nr = [], 0, 0
    i = 0
    while i < len(lines):
        ln = lines[i]
        if ln.lstrip().startswith("```"):
            fence = not fence
            i += 1
            continue
        if not fence and ln.strip().startswith("|") and i + 1 < len(lines) and SEP.match(lines[i + 1]):
            ncol = len(split_row(ln)); nt += 1
            j = i + 2
            while j < len(lines) and lines[j].strip().startswith("|") and not lines[j].lstrip().startswith("```"):
                nr += 1
                if len(split_row(lines[j])) != ncol:
                    bad.append((j + 1, len(split_row(lines[j])), ncol))
                j += 1
            i = j
        else:
            i += 1
    return nt, nr, bad


print("P1 文档表格结构（列数自检，§8.0.6 F4）")
for name, p in (("SKILL.md", SKILL), ("附录B", APX_B), ("附录D", APX_D)):
    nt, nr, bad = check_tables(p)
    say("P1", not bad, "%s：%d 表 / %d 数据行，列数不符 %d 处%s"
        % (name, nt, nr, len(bad), (" -> " + str(bad[:5])) if bad else ""))

# ---------------- P2 版本串 ----------------
print("P2 版本串一致")
sk = open(SKILL, encoding="utf-8").read()
tb = open(TOOL, encoding="utf-8").read()
ad = open(APX_D, encoding="utf-8").read()


def _mm(v):
    """取前两段做比较：front-matter 写 semver（1.10.0），标题写 v1.10。"""
    return ".".join(v.lstrip("v").split(".")[:2])

v_title = re.search(r"^# DRT 拟合原则 (v[\d.]+)", sk, re.M)
v_fm = re.search(r"^version:\s*([\d.]+)", sk, re.M)
v_tool = re.search(r"^PROTOCOL_VERSION = \"(v[\d.]+)\"", tb, re.M)
v_prot = re.search(r"^PROTOCOL = \"[^\"]*?(v[\d.]+)", tb, re.M)
v_apxd = re.search(r"DRT 拟合原则 (v[\d.]+) —", ad)
say("P2", v_title and v_fm and v_tool and v_apxd
    and v_title.group(1) == v_tool.group(1) == v_apxd.group(1)
    and _mm(v_fm.group(1)) == _mm(v_title.group(1)),
    "SKILL 标题=%s / front-matter=%s / 工具箱 PROTOCOL_VERSION=%s（PROTOCOL 串内 %s）/ 附录D 页脚=%s"
    % (v_title.group(1), v_fm.group(1), v_tool.group(1), v_prot.group(1), v_apxd.group(1)))
# P2b（v1.11）：§0 的"用户模板"里也写着版本号 —— 升版时最容易漏改的就是这一行
# （v1.10→v1.11 时实测漏了：模板还写着 v1.9），模板是用户直接复制粘贴的东西，必须与当前版本一致。
tpl = re.search(r"【遵循】DRT 拟合原则 (v[\d.]+)", sk)
say("P2b", tpl and tpl.group(1) == v_title.group(1),
    "§0 用户模板【遵循】行 = %s（应 = 标题 %s）" % (tpl.group(1) if tpl else "缺失", v_title.group(1)))
# P2c（v1.14）：工具箱与附录 D 的**文件头声明**（docstring 第 2 行）也带版本号 ——
# v1.13 实测漏改：那一行仍写着"协议 v1.12"，而 PROTOCOL_VERSION 已是 v1.13；
# 旧 P2 只查唯一常量与附录页脚，漏掉了这一行，于是"版本串一致"是假 PASS。
h_tool = re.search(r"—— 协议 (v[\d.]+) §", tb)
h_apxd = re.search(r"—— 协议 (v[\d.]+) §", ad)
say("P2c", h_tool and h_apxd and h_tool.group(1) == v_tool.group(1) == h_apxd.group(1),
    "工具箱/附录D 文件头『协议 %s / %s』应 = PROTOCOL_VERSION %s"
    % (h_tool.group(1) if h_tool else "缺失", h_apxd.group(1) if h_apxd else "缺失", v_tool.group(1)))

# ---------------- P3 附录 D 逐字节对齐 ----------------
print("P3 附录 D = 工具箱（逐字节）")
a = ad.index("```python\n") + len("```python\n")
b = ad.rindex("\n```")
extracted = ad[a:b]
sha = hashlib.sha256(tb.encode("utf-8")).hexdigest()
say("P3", extracted.rstrip("\n") == tb.rstrip("\n"), "代码段一致（%d 字符）" % len(extracted))
declared_sha = re.search(r"`sha256 = ([0-9a-f]{64})`", ad).group(1)
declared_n = re.search(r"（(\d+) 字节，(\d+) 行）", ad)
say("P3", declared_sha == sha, "声明 sha256 == 实际（%s…）" % sha[:12])
say("P3", int(declared_n.group(1)) == os.path.getsize(TOOL)
    and int(declared_n.group(2)) == tb.count("\n") + (0 if tb.endswith("\n") else 1),
    "声明字节/行数 == 实际（%s 字节 / %s 行）" % (declared_n.group(1), declared_n.group(2)))

# ---------------- P4 common_window 参考实现 ----------------
print("P4 §2 的 common_window() 参考实现")


def common_window(f, i_star):
    f_valley = [f[i] for i in i_star]
    f_min_common = max(f_valley)
    j = min((i for i, fv in zip(i_star, f_valley) if fv == f_min_common), default=min(i_star))
    return j, f_min_common


f_desc = [1e6 / (10 ** (k / 10.0)) for k in range(71)]          # 递减频率
# 合成样例：6 圈的 R2′ 谷点频率（Hz）—— 只用合成值，不含任何实测数据
freqs = [100.0, 50.0, 50.0, 25.0, 50.0, 50.0]
idx = [min(range(len(f_desc)), key=lambda i: abs(f_desc[i] - fr)) for fr in freqs]
j, fmc = common_window(f_desc, idx)
say("P4", abs(fmc - 100.0) < 0.5 and abs(f_desc[j] - 100.0) < 0.5,
    "f_min_common=%.6g Hz（索引 %d）—— 取到的是**最高**谷点，非最大索引 %d" % (fmc, j, max(idx)))
# 反向样例：两个谷点 12.5 / 25 Hz（本测试网格 10 点/十倍频，故容差取一个网格步长）
i2 = [min(range(len(f_desc)), key=lambda i: abs(f_desc[i] - fr)) for fr in (12.5, 25.0)]
j2, fmc2 = common_window(f_desc, i2)
step = f_desc[i2[1] - 1] - f_desc[i2[1]] if i2[1] > 0 else 0.1
say("P4", abs(fmc2 - 25.0) <= step and fmc2 > 20.0,
    "合成实例（12.5/25 Hz）-> %.6g Hz（网格步长 %.3g，取到的是两者中更高频者）" % (fmc2, step))

# ---------------- P5 工具箱可调用 ----------------
print("P5 工具箱子命令")
for sub in ("summary", "check", "normalize", "numbers", "verify"):
    r = subprocess.run([sys.executable, TOOL, sub, "--help"], capture_output=True, text=True,
                       encoding="utf-8")
    say("P5", r.returncode == 0, "%s --help 退出码 %d" % (sub, r.returncode))

# ---------------- P6 打包合规（v1.11 新增） ----------------
# 这一节回答"把这个包单独拿走/上传，还能不能用"：frontmatter 是否齐全、引用是否悬空、
# 编码是否会被解析器拒收、脚本是否还能编译。缺了它，包一旦脱离构建脚本就无人把关。
print("P6 打包合规（frontmatter / 引用完整性 / 编码 / 可编译）")
REQUIRED_FM = ["description", "description_zh", "description_en", "version", "author", "license"]
raw_skill = open(SKILL, "rb").read()
fm = {}
if raw_skill.startswith(b"---\n"):
    end = raw_skill.find(b"\n---\n", 3)
    if end > 0:
        for line in raw_skill[4:end].decode("utf-8").split("\n"):
            m = re.match(r"^([A-Za-z_]\w*):\s*(.*)$", line)
            if m:
                fm[m.group(1)] = m.group(2).strip().strip('"')
miss_fm = [k for k in REQUIRED_FM if not fm.get(k)]
say("P6", fm.get("name") == os.path.basename(BASE),
    "frontmatter name=%r == 包目录名 %r" % (fm.get("name"), os.path.basename(BASE)))
say("P6", not miss_fm, "必填字段齐全（%s）%s" % ("/".join(REQUIRED_FM),
                                                 ("  缺 " + str(miss_fm)) if miss_fm else ""))
# 许可与署名必须在包内（references/ 下），且含版权行与授权/原创性声明
_lic = os.path.join(BASE, "references", "许可与署名.md")
if os.path.exists(_lic):
    _lt = open(_lic, encoding="utf-8").read()
    say("P6", ("Apache License" in _lt) and ("Version 2.0, January 2004" in _lt)
        and ("Copyright 2026 Yangming Hu" in _lt) and ("原创性声明" in _lt)
        and ("PolyForm" not in _lt),
        "references/许可与署名.md 含 Apache-2.0 正文 + 版权行 + 原创性声明（无旧许可残留）")
else:
    say("P6", False, "references/许可与署名.md 缺失")
# 根目录结构：只允许 SKILL.md + references/ + scripts/ + templates/（WorkBuddy 技能结构）
_allowed_root = {"SKILL.md", "references", "scripts", "templates"}
_extra_root = sorted(x for x in os.listdir(BASE) if x not in _allowed_root)
say("P6", not _extra_root,
    "根目录只含 SKILL.md + references/ + scripts/ + templates/%s" %
    (("  多出 " + str(_extra_root)) if _extra_root else ""))

pkg_files = []
for dp, dns, fns in os.walk(BASE):
    dns[:] = [d for d in dns if d != "__pycache__"]
    for fn in fns:
        pkg_files.append(os.path.relpath(os.path.join(dp, fn), BASE).replace("\\", "/"))
inside = set(pkg_files)
# 占位符/通配符（本脚本自己的注释里就写着 `@references/…`、`scripts/…`）不算引用
def _real(toks):
    return [t for t in toks if not any(c in t for c in "…*<>{}")]
refs = _real(sorted(set(re.findall(r"@references/([^\s`）)、,，。]+)", sk))))
bad_refs = [r for r in refs if "references/" + r not in inside]
say("P6", not bad_refs, "@references 引用 %d 个全部存在%s" % (len(refs), ("  缺 " + str(bad_refs)) if bad_refs else ""))
scr = _real(sorted(set(re.findall(r"scripts/([^\s`）)、,，。]+\.py)", sk))))
bad_scr = [x for x in scr if "scripts/" + x not in inside]
say("P6", not bad_scr, "正文引用的 %d 个脚本全部存在%s" % (len(scr), ("  缺 " + str(bad_scr)) if bad_scr else ""))

# 禁用模式**拼出来**而不是字面写：否则本脚本自己的源码会命中，形成假 FAIL
BS = chr(92)
FORBID = [r"[A-Za-z]:" + BS + BS + "Users", "%USER" + "PROFILE%", BS + BS + ".dsh" + BS + BS + "skills"]
bad_enc, bad_abs, bad_syn, bad_nul = [], [], [], []
for r in pkg_files:
    b = open(os.path.join(BASE, r), "rb").read()
    if b.startswith(b"\xef\xbb\xbf") or b"\r\n" in b:
        bad_enc.append(r)
    # v1.13：表格里的 `\|` 曾被误写成 NUL（v1.12 实测 7 处，渲染时文字断裂）。
    # 列数自检恰好测不出来（它先把 `\|` 换成 NUL 再切分），所以必须单独断言。
    if b"\x00" in b:
        bad_nul.append("%s(%d)" % (r, b.count(b"\x00")))
    t = b.decode("utf-8", "ignore")
    if any(re.search(p, t) for p in FORBID):
        bad_abs.append(r)
    if r.endswith(".py"):
        try:
            compile(t, r, "exec")            # 内存编译：不写 __pycache__
        except SyntaxError as e:
            bad_syn.append("%s: %s" % (r, e))
say("P6", not bad_enc, "全部 %d 个文件 UTF-8 无 BOM + 纯 LF%s" % (len(pkg_files),
                                                                  ("  违规 " + str(bad_enc)) if bad_enc else ""))
say("P6", not bad_nul, "包内无 NUL 字节（§8.0.6 F4）%s" % (("  违规 " + str(bad_nul)) if bad_nul else ""))
say("P6", not bad_abs, "无本机绝对路径残留%s" % (("  违规 " + str(bad_abs)) if bad_abs else ""))
say("P6", not bad_syn, "全部脚本内存编译通过%s" % (("  失败 " + str(bad_syn)) if bad_syn else ""))
# 注意：pkg_files 的 walk 已把 __pycache__ 剪掉（那些是二进制 .pyc，不该进编码检查），
# 所以"无 __pycache__ 残留"**必须单独扫文件系统** —— 若拿 pkg_files 去判，该断言恒为真，
# 只是一句装饰（§8.0.6 B4；v1.11 实测踩过：断言通过但包内确实躺着 scripts/__pycache__/*.pyc）。
pyc = []
for dp, dns, fns in os.walk(BASE):
    if os.path.basename(dp) == "__pycache__":
        pyc += [os.path.relpath(os.path.join(dp, f), BASE).replace("\\", "/") for f in fns]
say("P6", not pyc, "包内无 __pycache__ / .pyc 残留%s" % (("  残留 " + str(pyc[:5])) if pyc else ""))

print("\n协议包自检结论：", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
