# -*- coding: utf-8 -*-
"""工具箱回归测试（v1.10 新增）—— 不依赖任何真实数据，全合成、可重复。

固化 v1.10 的两处修复 + v1.9 的两条硬要求：
  T1 `numbers` 幂等：重复运行基准条目数恒定（G1：不再扫自己的 dump）
  T2 `summary` 总是产出 §8.3 `DRT_batch_summary.csv`（17 列、BOM、LF、%.12g），
     且未加 `--csv` 时**不**产出矩阵 csv（§8.4 只出一个 xlsx）
  T3 `check` 对新建 workbook 全 PASS（H14 + H19）
  T4 `verify`（H21）对含必报数值的报告 PASS；把某个 `R_ohm` 改错后**必须** FAIL

用法：python scripts/自检_工具箱回归.py
退出码：0 = 全部 PASS；1 = 有 FAIL。
"""
import os
import re
import sys
import csv
import json
import shutil
import hashlib
import tempfile
import subprocess
import numpy as np

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOL = os.path.join(BASE, "scripts", "DRT归档工具箱.py")
ok = True


def say(tag, cond, msg):
    global ok
    ok = ok and bool(cond)
    print("  [%s] %s" % ("PASS" if cond else "FAIL", msg))


def run(*args):
    r = subprocess.run([sys.executable, TOOL] + list(args), capture_output=True, text=True,
                       encoding="utf-8")
    return r.returncode, (r.stdout or "") + (r.stderr or "")


root = tempfile.mkdtemp(prefix="drt_toolbox_selftest_")
D2 = os.path.join(root, "02_拟合结果_DRT格式")
D3 = os.path.join(root, "03_残差表")
D4 = os.path.join(root, "04_汇总表")
W = os.path.join(root, "_工作文件")
D7 = os.path.join(root, "07_报告")
for d in (D2, D3, D4, W, D7):
    os.makedirs(d)

# ---- 合成 3 个样品：单峰 DRT；残差为构造值（H7 判据内） ----
SAMPLES = [("synth A.txt", 20.0, 1.0e-6, 12.0), ("synth B.txt", 21.5, 1.2e-6, 9.0),
           ("synth C.txt", 22.5, 0.8e-6, 15.0)]
summary_rows = {}
for name, R, L, Rp in SAMPLES:
    tau = np.logspace(np.log10(1.6e-7), np.log10(1.3e-3), 400)
    g = Rp * np.exp(-0.5 * (np.log10(tau / 2.5e-5) / 0.35) ** 2)
    with open(os.path.join(D2, name.replace(".txt", "_fitted_DRT.txt")), "w",
              encoding="utf-8-sig", newline="\n") as fh:
        fh.write("L,%.12g\nR,%.12g\ntau,gamma\n" % (L, R))
        for t, gg in zip(tau, g):
            fh.write("%.12g,%.12g\n" % (t, gg))
    fs = np.logspace(6, np.log10(100.0), 40)
    zre = R + 8.0 / (1 + (2 * np.pi * fs * 3e-5) ** 2)
    zim = 2 * np.pi * fs * L - 8.0 * (2 * np.pi * fs * 3e-5) / (1 + (2 * np.pi * fs * 3e-5) ** 2)
    with open(os.path.join(D3, name.replace(".txt", "_residuals.csv")), "w",
              encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, lineterminator="\n")
        w.writerow(["f_Hz", "Zre_exp", "Zim_exp", "Zre_fit", "Zim_fit", "dRe_percent", "dIm_percent"])
        for i in range(len(fs)):
            w.writerow(["%.12g" % fs[i], "%.12g" % zre[i], "%.12g" % zim[i],
                        "%.12g" % (zre[i] * 1.001), "%.12g" % (zim[i] * 0.999),
                        "%.12g" % 0.1, "%.12g" % -0.1])

# ---- T1 numbers 幂等 ----
print("T1 numbers 幂等（G1）")
run("numbers", root)
p = os.path.join(W, "_report_numbers.json")
n1 = len(json.load(open(p, encoding="utf-8")))
run("numbers", root)
n2 = len(json.load(open(p, encoding="utf-8")))
say("T1", n1 == n2 > 0, "两次运行条目数 %d == %d" % (n1, n2))

# ---- T2 summary 产出 §8.3 csv ----
print("T2 summary -> §8.3 DRT_batch_summary.csv（G4）")
c, o = run("summary", D2, "EIS-DRT synth", "--out-dir", D4, "--note", "回归测试")
csvp = os.path.join(D4, "DRT_batch_summary.csv")
say("T2", c == 0 and os.path.exists(csvp), "summary 退出码 %d，csv 已生成" % c)
raw = open(csvp, "rb").read()
rows = list(csv.reader(open(csvp, encoding="utf-8-sig")))
want = ["sample", "cycle_index", "n_points", "f_min_Hz", "f_max_Hz", "R_ohm", "L_H",
        "R_pol_total_ohm", "R_pol_in_window_ohm", "R_pol_outside_window_ohm", "max_dRe_percent",
        "max_dIm_percent", "rms_percent", "min_gamma", "edge_fast_gamma", "edge_slow_gamma",
        "R_le_minZprime"]
say("T2", rows[0] == want and len(rows) == 4, "17 列 / 列序一致，%d 个数据行" % (len(rows) - 1))
say("T2", raw[:3] == b"\xef\xbb\xbf" and raw.count(b"\r\n") == 0, "UTF-8 BOM + 纯 LF")
say("T2", not os.path.exists(os.path.join(D4, "EIS-DRT synth.csv")), "未加 --csv 时无矩阵 csv（§8.4）")
summary_rows = {r[0]: r for r in rows[1:]}
say("T2", all(v[16] == "YES" for v in summary_rows.values()),
    "R_le_minZprime 真算（合成数据 R < min Z' → 全 YES）")

# ---- T3 check ----
print("T3 check（H14 + H19）")
c, o = run("check", os.path.join(D4, "EIS-DRT synth.xlsx"), "--dir", D2)
say("T3", c == 0 and "H14/H19 PASS" in o, "退出码 %d；%s" % (c, [l for l in o.splitlines() if "RESULT" in l]))

# ---- T4 verify（H21）：先写对，再故意写错 ----
print("T4 verify（H21）必报数值核对")
rep = os.path.join(D7, "DRT拟合报告_synth.md")
with open(rep, "w", encoding="utf-8", newline="\n") as fh:
    fh.write("# 回归测试报告\n\n| sample | R_ohm | R_pol_in_window_ohm |\n|---|---|---|\n")
    for name, *_ in SAMPLES:
        r = summary_rows[name]
        fh.write("| %s | %s | %s |\n" % (name, r[5], r[8]))
c1, o1 = run("verify", root)
say("T4", c1 == 0 and "H21 PASS" in o1, "报告写对 → 退出码 %d（%s）" % (c1, "PASS" if c1 == 0 else "FAIL"))
txt = open(rep, encoding="utf-8").read()
bad = txt.replace(summary_rows[SAMPLES[0][0]][5], "999.999999", 1)
open(rep, "w", encoding="utf-8", newline="\n").write(bad)
c2, o2 = run("verify", root)
miss = [l for l in o2.splitlines() if "H21" in l]
say("T4", c2 == 1 and "0" not in miss[0].split(":")[-1].strip(), "把第 1 个 R_ohm 改成 999.999999 → 退出码 %d（%s）"
    % (c2, miss[0].strip() if miss else "no H21 line"))
open(rep, "w", encoding="utf-8", newline="\n").write(txt)      # 复原

shutil.rmtree(root, ignore_errors=True)
print("\n工具箱回归测试结论：", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
