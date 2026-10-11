"""
【已被 `DRT归档工具箱.py` 取代 · v1.7】新批次请用工具箱的 `check` 子命令：
    python DRT归档工具箱.py check "<根>/04_汇总表/EIS-DRT <STEM>.xlsx" --dir "<根>/02_拟合结果_DRT格式"
本文件与工具箱功能等价（已实测：同一 workbook 结论一致），保留仅供复现历史批次；请勿在此就地修改。

H14 / H19 校验 —— 汇总表 EIS-DRT*.xlsx 与逐样品拟合文件的一致性（支持 union / intersection 轴）

用法:
  python DRT汇总表_H14校验.py <汇总表.xlsx> [--dir <拟合文件目录>]

检查项:
  1. 表头第 1 格恰为 tau；列数 = 数据集数 + 1
  2. tau 列严格单调递增
  3. 逐列：有值的单元格必须落在该数据集自己的测量窗口内，且 γ 与拟合文件一致
     （同网格直取路径 maxdiff = 0；插值路径容差 1e-9·max|γ|，见 §8.4 H14 第 1 条）
  4. 逐列：窗口内的每一行都必须有值 —— 不允许把有数据的点漏成空白
  5. 空白单元格绝不能被写成 0
  6. 【H19】缺残差表不得静默：窗口无法确定时判 FAIL（旧版会退回拟合网格，
     把外推边距当成测量窗口，H14 第 2/3 项静默失效）
  7. 【H19】Summary 页必须存在且遵守 §8.3 固定列名/列序；布尔列必须真算（YES/NO）
  8. 【H19】Info 页必须记录 protocol / axis mode / filled-blank，且计数与 DRT 页自洽，
     protocol 版本必须等于本文校验器的 PROTOCOL_VERSION
"""
import sys, os, glob, argparse
import numpy as np
from openpyxl import load_workbook

# H19：与协议文档、生成脚本同步升版（唯一来源，禁止在别处硬编码）
PROTOCOL_VERSION = "v1.6"

# §8.3 固定列（名字与顺序不得增删）——Summary 页前 17 列必须与之逐字一致，
# 之后允许追加（如 tau_window_min_s / tau_window_max_s）。
SUMMARY_SCHEMA = [
    "sample", "cycle_index", "n_points", "f_min_Hz", "f_max_Hz", "R_ohm", "L_H",
    "R_pol_total_ohm", "R_pol_in_window_ohm", "R_pol_outside_window_ohm",
    "max_dRe_percent", "max_dIm_percent", "rms_percent", "min_gamma",
    "edge_fast_gamma", "edge_slow_gamma", "R_le_minZprime",
]
# §8.0.4 H18 第 4 条：布尔列只允许 YES / NO（缺失写 NA）
BOOL_COLS = ["R_le_minZprime"]

# H14 第 2/3 项的窗口边界判据。校验器与生成器**必须逐字相同**：
# 不一致会让并集轴端点被误判成 GAP，产生整表假失败（实测曾整表所有列一起挂）。
BOUND_TOL = 1e-12


def read_fitted(path):
    tau = []; gam = []
    for line in open(path, encoding="utf-8-sig"):
        s = line.strip()
        if not s or s.startswith("L,") or s.startswith("R,") or s.lower().startswith("tau"):
            continue
        a, b = s.split(","); tau.append(float(a)); gam.append(float(b))
    tau = np.asarray(tau); gam = np.asarray(gam)
    o = np.argsort(tau)
    return tau[o], gam[o]


def residuals_path(src_fitted):
    """配套残差表的位置（同目录 或 03_残差表/）；找不到返回 None。"""
    base = os.path.basename(src_fitted).replace("_fitted_DRT.txt", "_residuals.csv")
    cands = [os.path.join(os.path.dirname(src_fitted), base),
             os.path.join(os.path.dirname(os.path.normpath(os.path.dirname(src_fitted))),
                          "03_残差表", base)]
    for c in cands:
        if os.path.exists(c):
            return c
    return None


def window_of(src_fitted):
    """measured window from the sibling / 03_残差表 residuals csv.

    找不到残差表时**不再静默退回拟合网格**：拟合网格含 0.5 dec 外推边距，
    退回后 H14 第 2/3 项（窗口内不漏空、窗口外必须空白）全部失效。
    """
    c = residuals_path(src_fitted)
    if c is not None:
        # §8.0.4 H18 的残差表是 UTF-8 带 BOM -> 必须 utf-8-sig，否则字段名带 '\ufeff'
        d = np.genfromtxt(c, delimiter=",", names=True, encoding="utf-8-sig")
        return (1/(2*np.pi*float(d["f_Hz"].max())),
                1/(2*np.pi*float(d["f_Hz"].min())), "residuals csv")
    t, _ = read_fitted(src_fitted)
    return t.min(), t.max(), "fitted grid (缺残差表 -> 窗口含外推边距，H14 第 2/3 项不可校验)"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--dir", default=None)
    a = ap.parse_args()

    d = a.dir or os.path.dirname(os.path.abspath(a.xlsx))
    wb = load_workbook(a.xlsx, data_only=True)
    fails, warns = [], []

    if "DRT" not in wb.sheetnames:
        print("[FAIL] no 'DRT' sheet; sheets =", wb.sheetnames); return 1
    ws = wb["DRT"]
    nrow, ncol = ws.max_row, ws.max_column
    hdr = [ws.cell(row=1, column=c).value for c in range(1, ncol + 1)]
    names = hdr[1:]

    ok = (hdr[0] == "tau")
    print("[%s] header[0] == 'tau'                 -> %r" % ("PASS" if ok else "FAIL", hdr[0]))
    if not ok: fails.append("header[0]")

    # 旧版的 `ncol != len(names) + 1` 恒为假（names 就是 hdr[1:]），属结构性空转的自检。
    # 真正要防的是：数据集列名缺失/重复 —— 列名是回查源文件的唯一钥匙。
    bad_names = [n for n in names if n is None or str(n).strip() == ""]
    dup_names = sorted({n for n in names if names.count(n) > 1})
    if bad_names or dup_names or not names:
        print("[FAIL] 表头数据集列名异常：空=%d 重复=%s" % (len(bad_names), dup_names))
        fails.append("header names")
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
            print("  [FAIL] %-46s source-missing（该列无法校验 -> 判 FAIL）" % str(name)[:46])
            fails.append("source-missing:" + str(name)); continue
        t, g = read_fitted(src)
        lo, hi, how = window_of(src)
        if how != "residuals csv":
            # H19：不得静默降级。旧版在这里退回拟合网格且不打印任何提示，
            # 于是"窗口外必须空白"变成对含外推边距的网格做检查 —— 门槛名存实亡。
            print("  [FAIL] %-46s 缺残差表 -> 窗口退化为拟合网格（含外推边距），"
                  "H14 第 2/3 项不可校验" % str(name)[:46])
            fails.append("no-residuals:" + str(name))
        col = [ws.cell(row=r, column=c).value for r in range(2, nrow + 1)]
        has = np.array([v is not None for v in col])
        vals = np.array([float(v) if v is not None else np.nan for v in col])
        nblank_total += int((~has).sum())

        # 3. value check on filled cells（插值路径允许浮点舍入，见 §8.4 H14）
        dmax = 0.0
        if has.any():
            ref = np.interp(np.log(tau_x[has]), np.log(t), g)
            dmax = float(np.max(np.abs(vals[has] - ref)))
            worst = max(worst, dmax)
        scale = float(np.nanmax(np.abs(vals[has]))) if has.any() else 1.0
        vtol = 1e-9 * max(scale, 1e-30)
        # 4. window coverage: all rows inside the window must be filled（判据与生成器一致）
        TOL = BOUND_TOL
        inside = (tau_x >= lo*(1 - TOL)) & (tau_x <= hi*(1 + TOL))
        missing = int((inside & ~has).sum())
        # out-of-window cells must be blank
        outside_filled = int(((~inside) & has).sum())

        flags = []
        if dmax > vtol: flags.append("VALUE")
        if missing: flags.append("GAP")
        if outside_filled: flags.append("OUT-OF-WINDOW")
        tag = "PASS" if not flags else "FAIL"
        vtxt = ("%.3g" % dmax) if dmax > 0 else "0"
        if 0 < dmax <= vtol:
            vtxt += " (~round-off, tol %.1e)" % vtol
        print("  [%s] %-46s filled=%3d blank=%3d win=[%.4g,%.4g]s maxdiff=%s%s"
              % (tag, str(name)[:46], int(has.sum()), int((~has).sum()), lo, hi, vtxt,
                 ("  <-- " + ",".join(flags)) if flags else ""))
        if missing: fails.append("gap:" + str(name))
        if outside_filled: fails.append("out-of-window:" + str(name))
        if dmax > vtol: fails.append("value:" + str(name))

    print("\n[%s] H14 worst absolute difference = %.3g" % ("PASS" if worst == 0 else "OK (round-off)", worst))
    fails = [f for f in fails if not f.startswith("value:") or worst > 1e-9]
    print("[INFO] blank cells in total = %d  (blank = that dataset has no valid value there)"
          % nblank_total)

    # ------------------------------------------------------------------
    # H19：Summary 页与 Info 页同样要校验。旧版只看 DRT 页，于是交付物里
    # 「Summary 首列叫 file、布尔列恒为 NO、Info 写着过期的协议版本」可以一路通过。
    # ------------------------------------------------------------------
    print("\n[H19] Summary / Info sheet checks:")
    if "Summary" not in wb.sheetnames:
        print("  [FAIL] 缺 Summary 工作表"); fails.append("no-summary-sheet")
        sm_rows, sm_hdr = [], []
    else:
        ws2 = wb["Summary"]
        sm_hdr = [ws2.cell(row=1, column=c).value for c in range(1, ws2.max_column + 1)]
        if sm_hdr[:len(SUMMARY_SCHEMA)] != SUMMARY_SCHEMA:
            print("  [FAIL] Summary 列名/列序不符 §8.3")
            print("         期望前 %d 列: %s" % (len(SUMMARY_SCHEMA), SUMMARY_SCHEMA))
            print("         实际        : %s" % sm_hdr)
            fails.append("summary-schema")
        else:
            print("  [PASS] Summary 前 %d 列逐字符合 §8.3" % len(SUMMARY_SCHEMA))
        sm_rows = [[ws2.cell(row=r, column=c).value for c in range(1, len(sm_hdr) + 1)]
                   for r in range(2, ws2.max_row + 1)]
        if len(sm_rows) != len(names):
            print("  [FAIL] Summary 行数 %d != 数据集数 %d" % (len(sm_rows), len(names)))
            fails.append("summary-rows")
        if sm_hdr and sm_hdr[0] and sm_rows:
            got = [r[0] for r in sm_rows]
            if got != list(names):
                print("  [FAIL] Summary 的 sample 列与 DRT 页列名不一致（顺序或取值）")
                fails.append("summary-sample-mismatch")
            else:
                print("  [PASS] Summary 的 sample 列与 DRT 页列名逐项一致")
        # 布尔列：只许 YES / NO / NA
        for bc in BOOL_COLS:
            if bc not in sm_hdr:
                continue
            j = sm_hdr.index(bc)
            bad = sorted({str(r[j]) for r in sm_rows if str(r[j]) not in ("YES", "NO", "NA")})
            if bad:
                print("  [FAIL] %s 列取值不合 H18 第 4 条（只许 YES/NO/NA）: %s" % (bc, bad))
                fails.append("bool:" + bc)
            counts = {}
            for r in sm_rows:
                counts[str(r[j])] = counts.get(str(r[j]), 0) + 1
            print("  [INFO] %s 取值分布: %s" % (bc, counts))
            # 真算校验：R_le_minZprime 必须与实测 min Z' 自洽（旧版写死 R<=1 -> 恒 NO）
            if bc == "R_le_minZprime" and "R_ohm" in sm_hdr and sm_rows:
                iR = sm_hdr.index("R_ohm")
                mismatch = []
                for r in sm_rows:
                    nm = r[0]
                    src = os.path.join(d, str(nm).replace(".txt", "_fitted_DRT.txt"))
                    res_path = residuals_path(src)
                    if res_path is None or not os.path.exists(src):
                        continue
                    rr = np.genfromtxt(res_path, delimiter=",", names=True, encoding="utf-8-sig")
                    if "Zre_exp" not in (rr.dtype.names or ()):
                        continue
                    rv = r[iR]
                    if rv in (None, "NA"):
                        continue
                    expect = "YES" if float(rv) <= float(np.min(rr["Zre_exp"])) else "NO"
                    if str(r[j]) != expect:
                        mismatch.append("%s(写 %s, 应为 %s, R=%s, minZ'=%.6g)"
                                        % (nm, r[j], expect, rv, float(np.min(rr["Zre_exp"]))))
                if mismatch:
                    print("  [FAIL] R_le_minZprime 与 H4 实际结论不符: %s" % "; ".join(mismatch[:3]))
                    fails.append("h4-column-wrong")
                else:
                    print("  [PASS] R_le_minZprime 与 H4（R ≤ min Z'）逐行自洽")

    if "Info" not in wb.sheetnames:
        print("  [FAIL] 缺 Info 工作表（协议版本/轴模式/填充计数无处溯源）")
        fails.append("no-info-sheet")
    else:
        wi = wb["Info"]
        info = {}
        for r in range(1, wi.max_row + 1):
            k = wi.cell(row=r, column=1).value
            if k:
                info[str(k).strip().lower()] = wi.cell(row=r, column=2).value
        for need in ("protocol", "axis mode", "filled / blank"):
            if need not in info:
                print("  [FAIL] Info 缺字段 '%s'" % need); fails.append("info:" + need)
        proto = str(info.get("protocol", ""))
        if PROTOCOL_VERSION not in proto:
            # 同版要求：本校验器与产物必须由同一版协议产出。对照旧归档物时，
            # 该 FAIL 的含义是「此产物产自旧版协议」，属提示性失败 —— 必须在报告中
            # 声明其适用版本，不得让它冒充当前版结果。
            print("  [FAIL] Info 的 protocol=%r != 校验器版本 %s"
                  "（同版要求；若是旧归档物，须在报告中声明其适用版本）"
                  % (proto, PROTOCOL_VERSION))
            fails.append("stale-protocol")
        else:
            print("  [PASS] Info protocol = %s" % proto)
        mode = info.get("axis mode")
        print("  [INFO] axis mode recorded in the workbook: %s" % mode)
        fb = str(info.get("filled / blank", ""))
        try:
            nfilled, nblank = int(fb.split("filled")[0].strip()), int(
                fb.split("filled")[1].split(",")[1].split()[0])
            if nblank != nblank_total:
                print("  [FAIL] Info 记的空白数 %d != DRT 页实测 %d" % (nblank, nblank_total))
                fails.append("info-fillcount")
            else:
                print("  [PASS] Info 的 filled/blank 计数与 DRT 页自洽 (%d filled, %d blank)"
                      % (nfilled, nblank))
        except (ValueError, IndexError):
            print("  [WARN] 无法解析 Info 的 filled / blank: %r" % fb)

    print("\nRESULT:", "H14/H19 PASS" if not fails else ("H14/H19 FAIL: " + ", ".join(fails)))
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
