## 附录 D · 归档工具箱（`DRT归档工具箱.py`）

> **这一节是自包含备份，日常不要读它**（约 764 行）：直接运行与本文档同目录的 `DRT归档工具箱.py` 即可（§8.0.7 的命令）。只有换机器、或工具文件丢失时，才把下面整段复制成该文件名。
>
> 本附录与同目录文件**逐字节相同** —— `DRT归档工具箱.py` 的 `sha256 = f43f9a12845d92716ac4192deee3e6dd4704ce75349090da928a8bddd1c8d09f`（43133 字节，840 行）。改动工具后必须同步本附录，并重新验证两者一致。
>
> 依赖：`numpy` + `openpyxl`。五个子命令：`normalize`（§8.0 重排，默认 dry-run）、`summary`（§8.4 汇总表）、`check`（H14 + H19）、`numbers`（H21 基准数值清单）、`verify`（H21 完整性）。

```python
"""
DRT 归档工具箱 —— 协议 v1.20 §8.0 / §8.4 / §8.0.6（硬门槛 H14 / H16 / H17 / H18 / H19 / H20 / H21）
====================================================================================
单文件、只依赖 numpy + openpyxl，五个子命令：

  summary   生成 `EIS-DRT <STEM>.xlsx`（§8.4：首列 tau，其余列各数据集 γ；union 轴）
            并**总是**顺带写出 §8.3 的 `DRT_batch_summary.csv`（17 列，不含 tau_window_*；
            该文件是 §8.0.3 强制交付物，v1.10 起不再需要各批次自写导出脚本）。
  check     回读校验 DRT / Summary / Info 三页（H14 + H19）
  normalize 把【已有】的旧结果目录按 §8.0 重排（默认 dry-run，--apply 才动手）
  numbers   从产物 dump 权威数值清单（H21 的比对基准）
  verify    校验报告里"必报数值"是否出现且未写错（H21；全自动，不产生人工清单）

用法：
  python DRT归档工具箱.py summary   "<根>/02_拟合结果_DRT格式" "EIS-DRT <STEM>" \
                                    --out-dir "<根>/04_汇总表" --note "<一句话>"
  python DRT归档工具箱.py check     "<根>/04_汇总表/EIS-DRT <STEM>.xlsx" \
                                    --dir "<根>/02_拟合结果_DRT格式"
  python DRT归档工具箱.py normalize "<结果目录>" [--stem "<STEM>"] [--apply]
  python DRT归档工具箱.py numbers   "<根>"
  python DRT归档工具箱.py verify    "<根>"

三条最容易被踩的坑（详见协议 §8.0.6）：
  A1  §8.0.4 H18 要求 csv 带 BOM，读端必须显式 utf-8-sig，否则列名变成 '\ufefff_Hz'；
  A3  csv.writer 默认行终止符是 \r\n，写成 LF 必须 lineterminator="\\n"；
  B1  找不到残差表时**不得**退回拟合网格（含外推边距）—— 直接判 FAIL，否则 H14 第 2/3 项静默失效。

关键实现约束（都是实测踩过的坑，改代码时不要退化）：
  · `numbers` 必须跳过 `_report_numbers.json` 自身，否则基准条目数逐次膨胀（每跑一次就翻倍）；
  · `summary` **总是**由 Summary 页数据产出 §8.3 的 17 列 csv（各批次不得再自写导出脚本）；
  · `verify` 接受整数 token（`%.12g` 会把整数写成 `20`），但整数用**零容差**精确比对，
    否则报告里任意一个 `15` 都会被当成 `15.17`（假 PASS）；
  · STEM 自动规则只在**真前缀**时有效：单文件批次 / 同名批次返回 None，必须由用户 `--stem` 指定
    （v1.13 修订；旧版 `stem_of()` 是死代码、`normalize` 另有一份内联实现，单文件批次下会
    把样品级后缀（序号 / 批次尾巴）当成体系名）。
"""
import sys, os, re, csv, glob, argparse, hashlib, shutil
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")   # D3：控制台 GBK
except Exception:
    pass

PROTOCOL = "DRT 拟合原则 v1.20, §8.4（§8.0 命名与目录 H16/H17/H18；§8.0.6 H19/H20；H1′/H5′/H7′；v1.20：同组 L 恒定判据、切弧三判据、峰突出度；许可 Apache-2.0）"
PROTOCOL_VERSION = "v1.20"         # H19：与协议文档、本文件唯一常量同步升版
BOUND_TOL = 1e-12                  # H14 第 6 条：生成器与校验器共用同一边界判据
SUMMARY_SCHEMA = ["sample", "cycle_index", "n_points", "f_min_Hz", "f_max_Hz", "R_ohm", "L_H",
                  "R_pol_total_ohm", "R_pol_in_window_ohm", "R_pol_outside_window_ohm",
                  "max_dRe_percent", "max_dIm_percent", "rms_percent", "min_gamma",
                  "edge_fast_gamma", "edge_slow_gamma", "R_le_minZprime"]
CANON_DIRS = ["01_去尾后数据", "02_拟合结果_DRT格式", "03_残差表", "04_汇总表",
              "05_图", "06_脚本", "07_报告", "_工作文件"]
FIG_MAP = {  # 旧图名 -> 规范图名（§8.0.2 H17）
    "图1_DRT叠图与窗口放大.png": "图1_DRT叠图_{s}.png", "图1_DRT叠图.png": "图1_DRT叠图_{s}.png",
    "DRT_vs_tau.png": "图1_DRT叠图_{s}.png", "DRT_batch_overview.png": "图1_DRT叠图_{s}.png",
    "图2_Nyquist_nyquist.png": "图2_Nyquist_{s}.png", "图2_Nyquist.png": "图2_Nyquist_{s}.png",
    "Nyquist_reconstructed.png": "图2_Nyquist_{s}.png",
    "图2_Nyquist_nyquist_highfreq.png": "图2b_Nyquist高频放大_{s}.png",
    "图2b_Nyquist高频放大.png": "图2b_Nyquist高频放大_{s}.png",
    "Nyquist_reconstructed_HF.png": "图2b_Nyquist高频放大_{s}.png",
    "图3_残差图.png": "图3_残差_{s}.png", "图3_残差.png": "图3_残差_{s}.png",
    "图3b_残差叠图.png": "图3b_残差叠图_{s}.png",
    "图4_跨圈趋势.png": "图4_参数趋势_{s}.png", "图4_参数趋势.png": "图4_参数趋势_{s}.png",
    "DRT_cycle_evolution.png": "图4_参数趋势_{s}.png"}
WORKFILE_RE = re.compile(r"^(_cache_.*\.npz|_results\.json|_log\.txt|.*\.log)$", re.I)
SAMPLE_SUFFIXES = ["_trimmed.txt", "_fitted_DRT.txt", "_residuals.csv",
                   "_DRT.txt", "_DRT.png", "_Nyquist.png", "_residuals.png"]


# ---------------------------------------------------------------- 公共工具
def md5(p):
    h = hashlib.md5()
    with open(p, "rb") as fh:
        for c in iter(lambda: fh.read(1 << 16), b""):
            h.update(c)
    return h.hexdigest()


def natural_key(p):
    n = re.findall(r"\d+", os.path.basename(p))
    return [int(x) for x in n] or [0]


def read_fitted(path):
    L = R = None; tau = []; gam = []
    for line in open(path, encoding="utf-8-sig"):          # A1/A2：读端也必须 utf-8-sig
        s = line.strip()
        if not s:
            continue
        if s.startswith("L,"):
            L = float(s.split(",")[1])
        elif s.startswith("R,"):
            R = float(s.split(",")[1])
        elif s.lower().startswith("tau"):
            continue
        else:
            a, b = s.split(","); tau.append(float(a)); gam.append(float(b))
    tau = np.asarray(tau); gam = np.asarray(gam)
    if np.any(np.diff(tau) <= 0):
        o = np.argsort(tau); tau, gam = tau[o], gam[o]
    return dict(name=os.path.basename(path).replace("_fitted_DRT.txt", ".txt"),
                tau=tau, gam=gam, R=R, L=L)


def residuals_path(src_fitted):
    base = os.path.basename(src_fitted).replace("_fitted_DRT.txt", "_residuals.csv")
    d = os.path.dirname(os.path.abspath(src_fitted))
    for c in (os.path.join(d, base), os.path.join(d, "03_残差表", base),
              os.path.join(os.path.dirname(d), "03_残差表", base)):
        if os.path.exists(c):
            return c
    return None


def window_of(src_fitted):
    """测量窗口 [tau(fmax), tau(fmin)]；**缺残差表不退回拟合网格**（B1）。"""
    c = residuals_path(src_fitted)
    if c is not None:
        x = np.genfromtxt(c, delimiter=",", names=True, encoding="utf-8-sig")   # A1
        return (1 / (2 * np.pi * float(x["f_Hz"].max())),
                1 / (2 * np.pi * float(x["f_Hz"].min())), "residuals csv")
    t, _ = read_fitted(src_fitted)
    return t.min(), t.max(), "fitted grid (缺残差表 -> 窗口含外推边距，H14 第 2/3 项不可校验)"


def _strip_eis(s):
    """§8.0.1：去掉开头的 'EIS '（大小写不敏感）—— 汇总表名已是 'EIS-DRT <STEM>'，不重复写 EIS。"""
    return re.sub(r"^EIS\s+", "", s, flags=re.I)


def common_prefix(names):
    if not names:
        return ""
    pre = os.path.splitext(names[0])[0]
    for n in names[1:]:
        s = os.path.splitext(n)[0]; i = 0
        while i < min(len(pre), len(s)) and pre[i] == s[i]:
            i += 1
        pre = pre[:i]
    return pre.rstrip(" _-#")   # 尾部序号分隔符（含 #）一并剥离


def stem_of(paths):
    """§8.0.1 STEM = 全体文件名去扩展名后的公共前缀，去尾部空格/`_`/`-`，去开头 'EIS '。

    v1.13：前缀必须是**真前缀** —— 至少两个文件的去扩展名名不同，且前缀不等于其中任何一个文件名。
    单文件批次 / 同名批次下自动规则无效，返回 None（由用户用 --stem 指定）；否则会把样品级后缀
    （如序号尾巴）当成体系名，产出 `EIS-DRT <单样品名>.xlsx` 这类违反 H17 溯源语义的文件名。"""
    names = [os.path.basename(p) for p in paths]
    stems = {os.path.splitext(n)[0] for n in names}
    if len(stems) < 2:
        return None
    pre = common_prefix(names)
    if not pre or pre in stems:
        return None
    return _strip_eis(pre)


# ---------------------------------------------------------------- summary
def build_axis(items, mode):
    los = [1 / (2 * np.pi * it["fmax"]) if it["fmin"] else it["tau"].min() for it in items]
    his = [1 / (2 * np.pi * it["fmin"]) if it["fmin"] else it["tau"].max() for it in items]
    lo = max(los) if mode == "intersection" else min(los)
    hi = min(his) if mode == "intersection" else max(his)
    if lo >= hi:
        raise SystemExit("ERROR: no common tau window (%s mode) -> check §2 H1" % mode)
    g0 = items[0]["tau"]
    identical = all(g.shape == g0.shape and np.allclose(g, g0, rtol=0, atol=0) for g in
                    (it["tau"] for it in items))
    if identical:
        axis = g0[(g0 >= lo) & (g0 <= hi)]
        note = ("%s axis over [%.6g, %.6g] s; all datasets share one grid -> "
                "its own grid points used, NO interpolation" % (mode, lo, hi))
        return axis, list(zip(los, his)), note
    dens = max(len(it["tau"]) / np.log10(it["tau"].max() / it["tau"].min()) for it in items)
    n = max(2, int(round(np.log10(hi / lo) * dens)))
    axis = np.logspace(np.log10(lo), np.log10(hi), n)
    return axis, list(zip(los, his)), ("%s axis over [%.6g, %.6g] s; %d points at the finest "
                                       "available density (%.1f pts/decade) -> linear interpolation "
                                       "in ln(tau)" % (mode, lo, hi, n, dens))


def build_matrix(items, axis, windows, tol=BOUND_TOL):
    data = np.full((len(axis), len(items)), np.nan)
    for i, (it, (lo, hi)) in enumerate(zip(items, windows)):
        inside = (axis >= lo * (1 - tol)) & (axis <= hi * (1 + tol))
        if not inside.any():
            continue
        v = np.interp(np.log(axis[inside]), np.log(it["tau"]), it["gam"], left=np.nan, right=np.nan)
        data[inside, i] = np.maximum(v, 0.0)
    return data


def cmd_summary(a):
    if a.files:
        paths = [os.path.join(a.dir, x.strip()) for x in a.files.split(",")]
    else:
        paths = sorted(glob.glob(os.path.join(a.dir, "*_fitted_DRT.txt")), key=natural_key)
    if not paths:
        raise SystemExit("no fitted files found in " + a.dir)
    stems = {re.sub(r"_fitted_DRT\.txt$", "", os.path.basename(p)) for p in paths}
    if len(stems) < 2:
        print("  WARN: 本批次只有 %d 个样品 -> §8.0.1 的 STEM 自动规则无效（单文件批次）。"
              % len(stems))
        print("        请确认前缀 %r 里的体系名由用户给出并在报告中写明来源，"
              "**不要**把样品级后缀（序号 / 批次尾巴）当体系名。" % a.prefix)
    out_dir = a.out_dir or a.dir
    os.makedirs(out_dir, exist_ok=True)      # 规范结构下 out_dir = 04_汇总表/，不存在则建
    items = []
    for p in paths:
        it = read_fitted(p)
        res = None
        c = residuals_path(p)
        if c is not None:
            res = np.genfromtxt(c, delimiter=",", names=True, encoding="utf-8-sig")
        if res is not None:
            it.update(fmin=float(res["f_Hz"].min()), fmax=float(res["f_Hz"].max()),
                      maxRe=float(np.abs(res["dRe_percent"]).max()),
                      maxIm=float(np.abs(res["dIm_percent"]).max()),
                      rms_pct=float(np.sqrt(np.mean(np.concatenate(
                          [res["dRe_percent"], res["dIm_percent"]]) ** 2))),
                      npt=int(len(res)),
                      minZre=float(np.min(res["Zre_exp"])) if "Zre_exp" in res.dtype.names else None)
        else:
            it.update(fmin=None, fmax=None, maxRe=float("nan"), maxIm=float("nan"),
                      rms_pct=float("nan"), npt=0, minZre=None)
        g, t = it["gam"], it["tau"]; dln = np.gradient(np.log(t))
        it["R_pol_total"] = float(np.sum(g * dln))
        win = np.ones_like(t, bool) if not it["fmin"] else \
            (t >= 1 / (2 * np.pi * it["fmax"])) & (t <= 1 / (2 * np.pi * it["fmin"]))
        it["R_pol_win"] = float(np.sum(g[win] * dln[win]))
        it["R_pol_out"] = it["R_pol_total"] - it["R_pol_win"]
        it["min_gamma"] = float(g.min()); it["edge_fast"] = float(g[0]); it["edge_slow"] = float(g[-1])
        items.append(it)

    axis, windows, note = build_axis(items, a.mode)
    data = build_matrix(items, axis, windows)
    M = len(items); filled = int(np.isfinite(data).sum())
    print("mode        : %s" % a.mode)
    print("axis        : %d points, tau %.6g .. %.6g s" % (len(axis), axis.min(), axis.max()))
    print("datasets    : %d" % M)
    print("assembly    : %s" % note)
    print("cells       : %d total, %d filled, %d blank (%.1f%% blank)"
          % (data.size, filled, data.size - filled, 100.0 * (data.size - filled) / data.size))

    if a.csv:                                   # A3：必须显式 LF
        p = os.path.join(out_dir, a.prefix + ".csv")
        with open(p, "w", encoding="utf-8-sig", newline="") as fh:
            w = csv.writer(fh, lineterminator="\n")
            w.writerow(["tau"] + [it["name"] for it in items])
            for r in range(len(axis)):
                w.writerow(["%.12g" % axis[r]] + ["" if not np.isfinite(v) else "%.12g" % v
                                                  for v in data[r]])
        print("saved", p)

    xlsx = os.path.join(out_dir, a.prefix + ".xlsx")
    wb = Workbook(); ws = wb.active; ws.title = "DRT"
    ws.append(["tau"] + [it["name"] for it in items])
    for c in range(1, M + 2):
        ws.cell(row=1, column=c).font = Font(bold=True)
        ws.cell(row=1, column=c).alignment = Alignment(horizontal="center", wrap_text=True)
    for r in range(len(axis)):
        ws.cell(row=r + 2, column=1, value=float(axis[r]))
        for c in range(M):
            if np.isfinite(data[r, c]):
                ws.cell(row=r + 2, column=c + 2, value=float(data[r, c]))
            # else: 留空 —— 绝不写 0（§8.4 第 4 条）
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["A"].number_format = "0.0000E+00"
    for c in range(M):
        L = get_column_letter(c + 2); ws.column_dimensions[L].width = 13
        for r in range(len(axis)):
            cc = ws.cell(row=r + 2, column=c + 2)
            if cc.value is not None:
                cc.number_format = "0.0000E+00"
    ws.freeze_panes = "B2"

    ws2 = wb.create_sheet("Summary")
    ws2.append(SUMMARY_SCHEMA + ["tau_window_min_s", "tau_window_max_s"])
    for c in range(1, len(SUMMARY_SCHEMA) + 3):
        ws2.cell(row=1, column=c).font = Font(bold=True)
    srows = []
    for i, it in enumerate(items):
        lo, hi = windows[i]
        h4 = ("YES" if it["R"] <= it["minZre"] else "NO") if (it["R"] is not None and
                                                              it["minZre"] is not None) else "NA"
        if h4 == "NO":
            print("  WARNING: H4 FAIL for %s: R=%.6g > min Z'=%.6g" % (it["name"], it["R"], it["minZre"]))
        srows.append([it["name"], i + 1, it["npt"],
                      float(it["fmin"]) if it["fmin"] else "NA",
                      float(it["fmax"]) if it["fmax"] else "NA",
                      it["R"], it["L"], it["R_pol_total"], it["R_pol_win"], it["R_pol_out"],
                      it["maxRe"], it["maxIm"], it["rms_pct"], it["min_gamma"],
                      it["edge_fast"], it["edge_slow"], h4, lo, hi])
        ws2.append(srows[-1])
    for c, wd in zip("ABCDEFGHIJKLMNOPQRS",
                     (46, 11, 10, 11, 11, 11, 12, 15, 17, 18, 15, 15, 12, 12, 15, 15, 16, 17, 17)):
        ws2.column_dimensions[c].width = wd
    ws2.freeze_panes = "A2"

    # §8.3：DRT_batch_summary.csv（v1.10 起由本工具直接产出；17 列，**不带** tau_window_*）
    # H18：列名列序与 §8.3 完全一致、缺失写 NA（不写 nan/空）、布尔 YES/NO、%.12g + BOM + LF。
    p = os.path.join(out_dir, "DRT_batch_summary.csv")
    _INT = {"cycle_index", "n_points"}; _STR = {"sample", "R_le_minZprime"}
    with open(p, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, lineterminator="\n")
        w.writerow(SUMMARY_SCHEMA)
        for row in srows:
            cells = []
            for k, v in zip(SUMMARY_SCHEMA, row[:len(SUMMARY_SCHEMA)]):
                if v is None or (isinstance(v, float) and v != v):
                    cells.append("NA")
                elif k in _INT:
                    cells.append(str(int(v)))
                elif k in _STR:
                    cells.append(str(v))
                else:
                    cells.append("%.12g" % float(v))
            w.writerow(cells)
    print("saved", p)

    ws3 = wb.create_sheet("Info")
    for row in [["protocol", PROTOCOL], ["axis mode", a.mode], ["assembly", note],
                ["tau axis", "%d points, %.6g .. %.6g s" % (len(axis), axis.min(), axis.max())],
                ["units", "tau in seconds, gamma in Ohm"], ["datasets", str(M)],
                ["filled / blank", "%d filled, %d blank" % (filled, data.size - filled)],
                ["note", a.note], ["", ""],
                ["BLANK CELLS", "A blank cell means that dataset has NO valid value at that tau"],
                ["", "(its measured window does not cover it). Blanks are NOT zeros."],
                ["WARNING", "tau outside a dataset's own measured window is EXTRAPOLATION."]]:
        ws3.append(row)
    ws3.column_dimensions["A"].width = 16; ws3.column_dimensions["B"].width = 96
    ws3["A1"].font = Font(bold=True)
    wb.save(xlsx)
    print("saved", xlsx)
    return 0


# ---------------------------------------------------------------- check（H14 + H19）
def _cells(l):
    s = l.strip(); s = s[1:] if s.startswith("|") else s
    return re.split(r"(?<!\\)\|", s[:-1] if s.endswith("|") else s)


def cmd_check(a):
    d = a.dir or os.path.dirname(os.path.abspath(a.xlsx))
    wb = load_workbook(a.xlsx, data_only=True)
    fails = []
    if "DRT" not in wb.sheetnames:
        print("[FAIL] no 'DRT' sheet; sheets =", wb.sheetnames); return 1
    ws = wb["DRT"]; nrow, ncol = ws.max_row, ws.max_column
    hdr = [ws.cell(row=1, column=c).value for c in range(1, ncol + 1)]
    names = hdr[1:]
    ok = (hdr[0] == "tau")
    print("[%s] header[0] == 'tau'                 -> %r" % ("PASS" if ok else "FAIL", hdr[0]))
    if not ok: fails.append("header[0]")
    if not names or any(n is None or str(n).strip() == "" for n in names) or \
            len(set(names)) != len(names):
        print("[FAIL] 表头数据集列名异常（空/重复）"); fails.append("header names")
    print("[INFO] datasets = %d, tau points = %d, cols = %d" % (len(names), nrow - 1, ncol))
    tau_x = np.array([ws.cell(row=r, column=1).value for r in range(2, nrow + 1)], float)
    mono = bool(np.all(np.diff(tau_x) > 0))
    print("[%s] tau strictly increasing           -> %d points, %.6g .. %.6g s"
          % ("PASS" if mono else "FAIL", len(tau_x), tau_x.min(), tau_x.max()))
    if not mono: fails.append("tau monotonic")

    print("\nper-column checks (value / coverage / blanks):")
    worst = 0.0; nblank_total = 0
    for c, name in enumerate(names, start=2):
        src = os.path.join(d, str(name).replace(".txt", "_fitted_DRT.txt"))
        if not os.path.exists(src):
            print("  [FAIL] %-46s source-missing（该列无法校验 -> FAIL）" % str(name)[:46])
            fails.append("source-missing:" + str(name)); continue
        t, g = read_fitted(src)["tau"], read_fitted(src)["gam"]
        lo, hi, how = window_of(src)
        if how != "residuals csv":
            print("  [FAIL] %-46s 缺残差表 -> 窗口退化为拟合网格（含外推边距），"
                  "H14 第 2/3 项不可校验" % str(name)[:46])
            fails.append("no-residuals:" + str(name))
        col = [ws.cell(row=r, column=c).value for r in range(2, nrow + 1)]
        has = np.array([v is not None for v in col])
        vals = np.array([float(v) if v is not None else np.nan for v in col])
        nblank_total += int((~has).sum())
        dmax = 0.0
        if has.any():
            ref = np.interp(np.log(tau_x[has]), np.log(t), g)
            dmax = float(np.max(np.abs(vals[has] - ref))); worst = max(worst, dmax)
        scale = float(np.nanmax(np.abs(vals[has]))) if has.any() else 1.0
        vtol = 1e-9 * max(scale, 1e-30)
        inside = (tau_x >= lo * (1 - BOUND_TOL)) & (tau_x <= hi * (1 + BOUND_TOL))
        missing = int((inside & ~has).sum()); outside = int(((~inside) & has).sum())
        flags = []
        if dmax > vtol: flags.append("VALUE")
        if missing: flags.append("GAP")
        if outside: flags.append("OUT-OF-WINDOW")
        vtxt = ("%.3g" % dmax) if dmax > 0 else "0"
        if 0 < dmax <= vtol:
            vtxt += " (~round-off, tol %.1e)" % vtol
        print("  [%s] %-46s filled=%3d blank=%3d win=[%.4g,%.4g]s maxdiff=%s%s"
              % ("PASS" if not flags else "FAIL", str(name)[:46], int(has.sum()), int((~has).sum()),
                 lo, hi, vtxt, ("  <-- " + ",".join(flags)) if flags else ""))
        for f in flags:
            fails.append("%s:%s" % (f.lower(), name))
    print("\n[%s] H14 worst absolute difference = %.3g"
          % ("PASS" if worst == 0 else "OK (round-off)", worst))
    fails = [f for f in fails if not f.startswith("value:") or worst > 1e-9]
    print("[INFO] blank cells in total = %d  (blank = that dataset has no valid value there)"
          % nblank_total)

    # ---- H19：Summary / Info 页 ----
    print("\n[H19] Summary / Info sheet checks:")
    if "Summary" not in wb.sheetnames:
        print("  [FAIL] 缺 Summary 工作表"); fails.append("no-summary-sheet"); sm_hdr, sm_rows = [], []
    else:
        ws2 = wb["Summary"]
        sm_hdr = [ws2.cell(row=1, column=c).value for c in range(1, ws2.max_column + 1)]
        if sm_hdr[:len(SUMMARY_SCHEMA)] != SUMMARY_SCHEMA:
            print("  [FAIL] Summary 列名/列序不符 §8.3\n         期望: %s\n         实际: %s"
                  % (SUMMARY_SCHEMA, sm_hdr)); fails.append("summary-schema")
        else:
            print("  [PASS] Summary 前 %d 列逐字符合 §8.3" % len(SUMMARY_SCHEMA))
        sm_rows = [[ws2.cell(row=r, column=c).value for c in range(1, len(sm_hdr) + 1)]
                   for r in range(2, ws2.max_row + 1)]
        if len(sm_rows) != len(names):
            print("  [FAIL] Summary 行数 %d != 数据集数 %d" % (len(sm_rows), len(names)))
            fails.append("summary-rows")
        if sm_hdr and sm_rows and [r[0] for r in sm_rows] != list(names):
            print("  [FAIL] Summary 的 sample 列与 DRT 页列名不一致"); fails.append("summary-sample")
        else:
            print("  [PASS] Summary 的 sample 列与 DRT 页列名逐项一致")
        if "R_le_minZprime" in sm_hdr and "R_ohm" in sm_hdr and sm_rows:
            j, iR = sm_hdr.index("R_le_minZprime"), sm_hdr.index("R_ohm")
            bad = sorted({str(r[j]) for r in sm_rows if str(r[j]) not in ("YES", "NO", "NA")})
            if bad:
                print("  [FAIL] R_le_minZprime 取值不合 H18 第 4 条: %s" % bad); fails.append("bool")
            print("  [INFO] R_le_minZprime 取值分布: %s"
                  % {v: sum(1 for r in sm_rows if str(r[j]) == v) for v in {str(r[j]) for r in sm_rows}})
            mm = []
            for r in sm_rows:
                src = os.path.join(d, str(r[0]).replace(".txt", "_fitted_DRT.txt"))
                c2 = residuals_path(src)
                if c2 is None or not os.path.exists(src) or r[iR] in (None, "NA"):
                    continue
                rr = np.genfromtxt(c2, delimiter=",", names=True, encoding="utf-8-sig")
                if "Zre_exp" not in (rr.dtype.names or ()):
                    continue
                exp = "YES" if float(r[iR]) <= float(np.min(rr["Zre_exp"])) else "NO"
                if str(r[j]) != exp:
                    mm.append("%s(写 %s 应为 %s, R=%s, minZ'=%.6g)"
                              % (r[0], r[j], exp, r[iR], float(np.min(rr["Zre_exp"]))))
            if mm:
                print("  [FAIL] R_le_minZprime 与 H4（R ≤ min Z'）不符: %s" % "; ".join(mm[:3]))
                fails.append("h4-column-wrong")
            else:
                print("  [PASS] R_le_minZprime 与 H4（R ≤ min Z'）逐行自洽")
    if "Info" not in wb.sheetnames:
        print("  [FAIL] 缺 Info 工作表"); fails.append("no-info-sheet")
    else:
        wi = wb["Info"]; info = {}
        for r in range(1, wi.max_row + 1):
            k = wi.cell(row=r, column=1).value
            if k:
                info[str(k).strip().lower()] = wi.cell(row=r, column=2).value
        for need in ("protocol", "axis mode", "filled / blank"):
            if need not in info:
                print("  [FAIL] Info 缺字段 '%s'" % need); fails.append("info:" + need)
        proto = str(info.get("protocol", ""))
        if PROTOCOL_VERSION not in proto:
            print("  [FAIL] Info 的 protocol=%r != 校验器版本 %s"
                  "（同版要求；若是旧归档物，须在报告中声明其适用版本）" % (proto, PROTOCOL_VERSION))
            fails.append("stale-protocol")
        else:
            print("  [PASS] Info protocol = %s" % proto)
        print("  [INFO] axis mode recorded in the workbook: %s" % info.get("axis mode"))
        fb = str(info.get("filled / blank", ""))
        try:
            nb = int(fb.split("filled")[1].split(",")[1].split()[0])
            if nb != nblank_total:
                print("  [FAIL] Info 记的空白数 %d != DRT 页实测 %d" % (nb, nblank_total))
                fails.append("info-fillcount")
            else:
                print("  [PASS] Info 的 filled/blank 计数与 DRT 页自洽 (%s)" % fb)
        except (ValueError, IndexError):
            print("  [WARN] 无法解析 Info 的 filled / blank: %r" % fb)
    print("\nRESULT:", "H14/H19 PASS" if not fails else ("H14/H19 FAIL: " + ", ".join(fails)))
    return 0 if not fails else 1


# ---------------------------------------------------------------- normalize（§8.0）
def _index_of(name):
    m = re.match(r"^(.*?)(\d+)([^\d]*)$", name)
    return (m.group(1), int(m.group(2)), m.group(3)) if m else None


def cmd_normalize(a):
    d = os.path.abspath(a.result_dir)
    if not os.path.isdir(d):
        print("ERROR: 结果目录不存在: %s" % d); return 2
    if a.original_dir:
        dirs = [a.original_dir]
    else:
        dirs = sorted(glob.glob(os.path.join(d, "_backup_*"))) or [os.path.dirname(d)]
    orig = {}
    for dd in dirs:
        if not os.path.isdir(dd):
            print("[WARN] 原始数据目录不存在，跳过: %s" % dd); continue
        for p in sorted(glob.glob(os.path.join(dd, "*.txt"))):
            n = os.path.basename(p)
            if any(n.endswith(s) for s in ("_trimmed.txt", "_fitted_DRT.txt")):
                continue
            r = _index_of(n)
            if not r:
                continue
            # 完整性打分：点数多者优先（数据无关，不依赖任何命名约定）
            try:
                sc = sum(1 for _ in open(p, encoding="utf-8-sig", errors="ignore"))
            except Exception:
                sc = 0
            if r[1] not in orig or sc > orig[r[1]][1]:
                orig[r[1]] = (n, sc)
    if not orig:
        print("ERROR: 找不到原始数据文件；请用 --original-dir 指定。"); return 2
    orig = {k: v[0] for k, v in orig.items()}
    stem = a.stem or stem_of(list(orig.values()))
    if not stem:
        print("ERROR: 无法自动确定 STEM（单文件批次 / 同名批次 / 无公共前缀），请用 --stem 指定。")
        return 2
    print("结果目录  : %s" % d)
    print("原始数据  : %s" % "; ".join(dirs))
    print("STEM      : %r   （如不对请用 --stem 覆盖）" % stem)
    print("原始文件  : %d 个；序号映射（推断，请核对）:" % len(orig))
    for k in sorted(orig):
        print("   %-4s -> %s" % (k, orig[k]))

    plan, unknown = [], []
    orig_set = set(orig.values())

    def add(src, dst, why):
        if os.path.abspath(src) != os.path.abspath(dst):
            plan.append((os.path.relpath(src, d), os.path.relpath(dst, d), why))

    for root, _dirs, files in os.walk(d):
        rel = os.path.relpath(root, d); top = rel.split(os.sep)[0]
        if top.startswith("_backup_"):
            continue                                    # 备份一个字节都不动
        in_canon = rel != "." and top in CANON_DIRS
        for n in sorted(files):
            src = os.path.join(root, n)
            if WORKFILE_RE.match(n):
                add(src, os.path.join(d, "_工作文件", n), "中间件"); continue
            if n.lower().endswith(".py"):
                add(src, os.path.join(d, "06_脚本", n), "脚本"); continue
            if n == "DRT_batch_summary.csv":
                add(src, os.path.join(d, "04_汇总表", n), "汇总 csv"); continue
            if n.startswith("EIS-DRT") and n.lower().endswith((".xlsx", ".csv")):
                add(src, os.path.join(d, "04_汇总表",
                                      "EIS-DRT %s%s" % (stem, os.path.splitext(n)[1])),
                    "汇总表名归一"); continue
            if n in FIG_MAP:
                add(src, os.path.join(d, "05_图", FIG_MAP[n].format(s=stem)), "图名归一"); continue
            if n.startswith("DRT拟合报告") and n.lower().endswith(".md"):
                add(src, os.path.join(d, "07_报告", "DRT拟合报告_%s.md" % stem), "报告"); continue
            hit = next((s for s in SAMPLE_SUFFIXES if n.endswith(s)
                        and (n[:-len(s)] + ".txt") in orig_set), None)
            if hit:
                sub = {"_trimmed.txt": ("01_去尾后数据", "_trimmed.txt"),
                       "_fitted_DRT.txt": ("02_拟合结果_DRT格式", "_fitted_DRT.txt"),
                       "_residuals.csv": ("03_残差表", "_residuals.csv")}.get(hit, ("05_图", hit))
                add(src, os.path.join(d, sub[0], n[:-len(hit)] + sub[1]), "归位（名字已合规）")
                continue
            m = re.match(r"^(?:cycle|sample)?[ _-]*0*(\d+)(.*)$", os.path.splitext(n)[0], re.I)
            if m and int(m.group(1)) in orig:
                S = os.path.splitext(orig[int(m.group(1))])[0]; rest = m.group(2).lower()
                if rest.endswith("_drt") or n.endswith("_DRT.txt"):
                    dst = os.path.join(d, "02_拟合结果_DRT格式", S + "_fitted_DRT.txt")
                elif "trim" in rest:
                    dst = os.path.join(d, "01_去尾后数据", S + "_trimmed.txt")
                elif "resid" in rest:
                    dst = os.path.join(d, "03_残差表", S + "_residuals.csv")
                else:
                    unknown.append((os.path.relpath(src, d), "疑似样品级但后缀无法判定")); continue
                add(src, dst, "溯源改名（%d -> %s）" % (int(m.group(1)), os.path.basename(dst)))
                continue
            if not in_canon:
                unknown.append((os.path.relpath(src, d), "未识别（H16：根目录不得有文件）"))

    by_dst, dup = {}, []
    for s, t, _w in plan:
        if t in by_dst:
            dup.append((by_dst[t], s, t))
        else:
            by_dst[t] = s
    conflicts = [(s, t) for s, t, _w in plan
                 if os.path.exists(os.path.join(d, t))
                 and md5(os.path.join(d, s)) != md5(os.path.join(d, t))]
    print("=" * 108)
    print("计划重排 %d 项：" % len(plan))
    for s, t, w in plan:
        print("  %-56s -> %-56s [%s]" % (s, t, w))
    if dup:
        print("\n目标重复 %d 组：" % len(dup))
        for a1, b1, t in dup:
            same = md5(os.path.join(d, a1)) == md5(os.path.join(d, b1))
            print("  %s + %s -> %s [%s]" % (a1, b1, t, "内容相同，去重" if same else "内容不同，需人工判断"))
    if conflicts:
        print("\n冲突 %d 项（目标已存在且内容不同，--apply 会跳过，绝不覆盖）：" % len(conflicts))
        for s, t in conflicts:
            print("  %s -> %s" % (s, t))
    if unknown:
        print("\n需人工判断 %d 项：" % len(unknown))
        for s, w in unknown:
            print("  %-56s [%s]" % (s, w))
    print("\n目标目录结构（§8.0.3 H16）：%s + _backup_<DATE>/" % " / ".join(CANON_DIRS))
    if not a.apply:
        print("\n[DRY-RUN] 未改动任何文件。确认计划后加 --apply 执行。")
        return 0
    moved = dedup = skipped = 0
    for c in CANON_DIRS:
        os.makedirs(os.path.join(d, c), exist_ok=True)
    done = set()
    for s, t, _w in plan:
        src, dst = os.path.join(d, s), os.path.join(d, t)
        if t in done:
            if os.path.exists(src) and os.path.exists(dst) and md5(src) == md5(dst):
                os.remove(src); dedup += 1
            else:
                print("[SKIP 重复目标且内容不同] %s" % s); skipped += 1
            continue
        if not os.path.exists(src):
            continue
        if os.path.exists(dst):
            if md5(src) == md5(dst):
                os.remove(src); dedup += 1
            else:
                print("[SKIP 冲突] %s" % s); skipped += 1
            continue
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.move(src, dst); done.add(t); moved += 1
    print("\n[APPLIED] 移动/改名 %d 项，去重删除 %d 项，跳过 %d 项。" % (moved, dedup, skipped))
    print("提醒：重命名后请重跑 `check`，确认没有整列被 FAIL。")
    return 0


# ---------------------------------------------------------------- numbers / verify（H21）
# G7（v1.10）：token 必须也接受**不带小数点**的整数写法 —— `%.12g` 会把恰好是整数的
# 产物值写成 `20`，旧正则（只认 `\d+\.\d+`）取不到它，于是"报告写对了"也会被判缺失。
NUM_RE = re.compile(r"-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?")


def collect_numbers(root):
    """从产物 dump 权威数值清单（H21 的比对基准）：04_汇总表 的全部 csv、_工作文件 的 json、
    拟合文件的 R/L、汇总表 Info 的计数。

    v1.10（G1）：**必须跳过 `_report_numbers.json` 自身** —— 它是本函数的 dump 目标，
    以前会被当成"产物"再扫一遍，使基准条目数逐次膨胀（每跑一次就把上一轮的 dump 再扫一遍）。
    条目数不稳定会让读者以为产物变了；跳过它之后重复运行结果恒定。"""
    vals = []
    for p in sorted(glob.glob(os.path.join(root, "04_汇总表", "*.csv"))):
        try:
            rows = list(csv.DictReader(open(p, encoding="utf-8-sig")))
        except Exception:
            continue
        for row in rows:
            for k, v in row.items():
                if v in (None, "", "NA") or k in ("sample",):
                    continue
                try:
                    vals.append((float(v), "%s:%s" % (os.path.basename(p), k)))
                except (TypeError, ValueError):
                    pass
    for p in sorted(glob.glob(os.path.join(root, "_工作文件", "*.json"))):
        if os.path.basename(p) == "_report_numbers.json":      # G1：不扫自己的 dump
            continue
        def walk(o, path=""):
            if isinstance(o, dict):
                for k, v in o.items():
                    walk(v, path + "/" + str(k))
            elif isinstance(o, list):
                for i, v in enumerate(o):
                    walk(v, path + "[%d]" % i)
            elif isinstance(o, (int, float)) and not isinstance(o, bool):
                vals.append((float(o), os.path.basename(p) + path))
        try:
            import json
            walk(json.load(open(p, encoding="utf-8")))
        except Exception:
            pass
    for f in sorted(glob.glob(os.path.join(root, "02_拟合结果_DRT格式", "*_fitted_DRT.txt"))):
        it = read_fitted(f)
        vals.append((it["R"], "%s/R" % it["name"]))
        vals.append((it["L"], "%s/L" % it["name"]))
    for x in sorted(glob.glob(os.path.join(root, "04_汇总表", "EIS-DRT*.xlsx"))):
        wb = load_workbook(x, data_only=True)
        for sh in wb.sheetnames:
            if sh == "DRT":
                continue
            ws = wb[sh]
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(row=r, column=c).value
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        vals.append((float(v), "%s!%s" % (os.path.basename(x), sh)))
    return vals


def _ulp(token):
    """报告里这个数字"最后一位"的量级：12.345 -> 1e-3；1.570e-07 -> 1e-10。
    用它做容差，才不会被"报告少写了两位"误判。

    G7（v1.10）：纯整数写法（如 `20`、`40`）返回 **0.0**，即要求**精确相等** ——
    若给整数容差 0.5，报告里随便一个 `15` 都会被当成 `15.17`（假 PASS），
    而 H21 的失败方向必须是"宁可漏报、不可错报"。"""
    m = re.match(r"(-?\d+)\.(\d+)(?:[eE]([-+]?\d+))?$", token)
    if m:
        dec = len(m.group(2))
        exp = int(m.group(3)) if m.group(3) else 0
        return (10.0 ** -dec) * (10.0 ** exp)
    if re.match(r"-?\d+(?:[eE][-+]?\d+)?$", token):
        return 0.0
    return None


def _norm(val):
    """归一化为 (尾数 ∈ [1,10), 指数)，用于"允许 ×10^n 记法"的比较（避免 10**k 溢出/下溢）。"""
    import math
    if val == 0:
        return (0.0, 0)
    e = int(math.floor(math.log10(abs(val))))
    den = 10.0 ** e
    if den == 0 or den == float("inf"):        # denormal / 极端量级：退化为原值比较
        return (val, 0)
    return (val / den, e)


def _matches(v, vals, ulp):
    """v 是否等于某个产物数值（允许只差整数个十倍频；容差 = 报告写法最后一位的一半）。"""
    if ulp is None:
        return False
    mv, ev = _norm(v)
    den = 10.0 ** ev
    tol = (0.5 * ulp / den * 1.000001) if den > 0 else 0.0
    for x, _src in vals:
        if x == 0:
            continue
        mx, _ex = _norm(x)
        if abs(mv - mx) <= tol:
            return True
    return False


def cmd_numbers(a):
    vals = collect_numbers(os.path.abspath(a.root))
    p = os.path.join(a.root, "_工作文件", "_report_numbers.json")
    os.makedirs(os.path.dirname(p), exist_ok=True)
    import json
    json.dump([{"value": v, "source": s} for v, s in vals], open(p, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("authoritative numbers: %d  -> %s" % (len(vals), p))
    return 0


def cmd_verify(a):
    """H21（**完全由代码判定，不产生人工复核任务**）：
    每个样品的 `R_ohm` 与 `R_pol_in_window_ohm` 必须出现在报告中，且不得写成与之不符的数字。
    派生量与上下文常数**不做校验**（无法与产物值自动区分，不为此设置人工清单）。
    退出码：0 = PASS；1 = 有必报数值缺失/写错。"""
    root = os.path.abspath(a.root)
    vals = collect_numbers(root)
    rep = a.report or (sorted(glob.glob(os.path.join(root, "07_报告", "DRT拟合报告_*.md"))) or [None])[0]
    if not rep or not os.path.exists(rep):
        print("[FAIL] 找不到报告（H21 无法校验）"); return 1
    text = open(rep, encoding="utf-8").read()
    toks = [(m.group(0), _ulp(m.group(0)), float(m.group(0))) for m in NUM_RE.finditer(text)]
    print("report  : %s" % os.path.basename(rep))
    print("baseline: %d 个产物数值；报告候选数字 %d 个" % (len(vals), len(toks)))
    csvp = os.path.join(root, "04_汇总表", "DRT_batch_summary.csv")
    if not os.path.exists(csvp):
        print("[FAIL] 缺 DRT_batch_summary.csv（H21 无法校验）"); return 1
    missing = []
    for row in csv.DictReader(open(csvp, encoding="utf-8-sig")):
        for col in ("R_ohm", "R_pol_in_window_ohm"):
            try:
                x = float(row[col])
            except (TypeError, ValueError):
                continue
            if not any(_matches(v, [(x, "")], ulp) for _s, ulp, v in toks):
                missing.append("%s/%s=%.6g" % (row.get("sample", "?"), col, x))
    print("[H21] 必报数值（每样品 R_ohm 与 R_pol_in_window_ohm）缺失或写错: %d" % len(missing))
    for m in missing[:20]:
        print("    [FAIL] %s" % m)
    if len(missing) > 20:
        print("    … 另有 %d 项" % (len(missing) - 20))
    print("\nRESULT:", "H21 PASS" if not missing else "H21 FAIL: 必报数值 %d 项缺失/写错" % len(missing))
    return 0 if not missing else 1


# ---------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser(prog="DRT归档工具箱", description="DRT 归档工具箱（协议 §8.0 / §8.4）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("summary", help="生成 EIS-DRT <STEM>.xlsx（§8.4）+ DRT_batch_summary.csv（§8.3）")
    s.add_argument("dir"); s.add_argument("prefix")
    s.add_argument("--files", default=None); s.add_argument("--note", default="")
    s.add_argument("--mode", choices=["union", "intersection"], default="union")
    s.add_argument("--out-dir", default=None)
    s.add_argument("--csv", action="store_true",
                   help="额外出矩阵 csv（EIS-DRT <STEM>.csv）；§8.4 默认只出 xlsx。"
                        "§8.3 的 DRT_batch_summary.csv 不受此开关影响，总是输出。")
    s.set_defaults(fn=cmd_summary)

    c = sub.add_parser("check", help="回读校验 DRT/Summary/Info（H14 + H19）")
    c.add_argument("xlsx"); c.add_argument("--dir", default=None)
    c.set_defaults(fn=cmd_check)

    n = sub.add_parser("normalize", help="按 §8.0 重排旧结果目录（默认 dry-run）")
    n.add_argument("result_dir")
    n.add_argument("--stem", default=None,
                   help="体系名；§8.0.1 自动规则无效时（单文件批次 / 同名批次 / 无公共前缀）必填")
    n.add_argument("--original-dir", default=None); n.add_argument("--apply", action="store_true")
    n.set_defaults(fn=cmd_normalize)

    m = sub.add_parser("numbers", help="从产物 dump 权威数值清单（H21 基准）")
    m.add_argument("root")
    m.set_defaults(fn=cmd_numbers)

    v = sub.add_parser("verify", help="H21：报告数字是否都能追溯到产物")
    v.add_argument("root"); v.add_argument("--report", default=None)
    v.set_defaults(fn=cmd_verify)

    a = ap.parse_args()
    return a.fn(a)


if __name__ == "__main__":
    sys.exit(main())

```

---

*DRT 拟合原则 v1.20 — 建议在每次分析开始时把本文件全文连同数据一并提供给 AI。*
*文件名不带版本号，正文始终为最新版。*
