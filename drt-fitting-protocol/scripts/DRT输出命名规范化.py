"""
【已被 `DRT归档工具箱.py` 取代 · v1.7】新批次请用工具箱的 `normalize` 子命令：
    python DRT归档工具箱.py normalize "<结果目录>"            # dry-run
    python DRT归档工具箱.py normalize "<结果目录>" --apply
本文件与工具箱功能等价（已实测：同一旧目录重排后**目录树逐文件一致**），保留仅供复现历史批次。

DRT 输出命名规范化（协议 §8.0，硬门槛 H16 / H17 / H18）
================================================================
把【已经产生】的旧结果目录，按 §8.0 重命名并重排到固定目录结构。
不改数据内容，只动文件名与位置。

安全默认：**dry-run** —— 只打印计划，不动任何文件。确认无误后加 --apply。

用法:
  python DRT输出命名规范化.py <结果目录> [--stem "体系名"] [--original-dir <原始数据目录>] [--apply]

  结果目录       例如 "<结果根目录>"（旧目录结构，按 §8.0 重排）
  --stem         体系名 STEM（§8.0.1）。省略则自动取原始文件名的最长公共前缀
                 （去掉结尾空格/下划线，并去掉开头的 "EIS "）
  --original-dir 原始数据目录。省略则依次尝试：
                  1) <结果目录>/_backup_*/         （§2 H2 的备份）
                  2) <结果目录>/ 的父目录
  原始数据目录**只读**：本脚本绝不改名、绝不写入、绝不删除其中的任何文件。

固定目录（§8.0.3 H16）:
  01_去尾后数据 / 02_拟合结果_DRT格式 / 03_残差表 / 04_汇总表 /
  05_图 / 06_脚本 / 07_报告 / _工作文件 / _backup_<DATE>

固定文件名（§8.0.2 H17）:
  样品级  <S>_trimmed.txt / <S>_fitted_DRT.txt / <S>_residuals.csv
  批次级  DRT_batch_summary.csv / EIS-DRT <STEM>.xlsx / 图1…图4_<STEM>.png / DRT拟合报告_<STEM>.md
  其中 S = 原始数据文件名去扩展名（逐字符保留）。

注意：本脚本对 "cycleNN -> 第 NN 个原始文件" 的映射是**推断**，会在开头打印整张
映射表；--apply 之前请人工核对。若映射不对，请先手工改名或用 --stem 指定体系名。
"""
import sys, os, re, glob, argparse, hashlib, shutil

try:                                     # 控制台按 UTF-8 打印，避免中文乱码
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

CANON_DIRS = ["01_去尾后数据", "02_拟合结果_DRT格式", "03_残差表",
              "04_汇总表", "05_图", "06_脚本", "07_报告", "_工作文件"]
# 旧图名 -> 规范图名（{stem} 为体系名）
FIG_MAP = {
    "图1_DRT叠图与窗口放大.png": "图1_DRT叠图_{stem}.png",
    "图1_DRT叠图.png": "图1_DRT叠图_{stem}.png",
    "DRT_vs_tau.png": "图1_DRT叠图_{stem}.png",
    "DRT_batch_overview.png": "图1_DRT叠图_{stem}.png",
    "图2_Nyquist_nyquist.png": "图2_Nyquist_{stem}.png",
    "图2_Nyquist.png": "图2_Nyquist_{stem}.png",
    "Nyquist_reconstructed.png": "图2_Nyquist_{stem}.png",
    "图2_Nyquist_nyquist_highfreq.png": "图2b_Nyquist高频放大_{stem}.png",
    "图2b_Nyquist高频放大.png": "图2b_Nyquist高频放大_{stem}.png",
    "Nyquist_reconstructed_HF.png": "图2b_Nyquist高频放大_{stem}.png",
    "图3_残差图.png": "图3_残差_{stem}.png",
    "图3_残差.png": "图3_残差_{stem}.png",
    "图3b_残差叠图.png": "图3b_残差叠图_{stem}.png",
    "图4_跨圈趋势.png": "图4_参数趋势_{stem}.png",
    "图4_参数趋势.png": "图4_参数趋势_{stem}.png",
    "DRT_cycle_evolution.png": "图4_参数趋势_{stem}.png",
}
WORKFILE_RE = re.compile(r"^(_cache_.*\.npz|_results\.json|_log\.txt|.*\.log)$", re.I)
SAMPLE_SUFFIXES = ["_trimmed.txt", "_fitted_DRT.txt", "_residuals.csv",
                   "_DRT.txt", "_DRT.png", "_Nyquist.png", "_residuals.png"]


def md5(p):
    h = hashlib.md5()
    with open(p, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def strip_eis(s):
    """§8.0.1：STEM 去掉开头的 'EIS '（汇总表名已是 EIS-DRT <STEM>，不重复写）"""
    return re.sub(r"^EIS\s+", "", s, flags=re.I)


def is_data_file(name):
    """原始三列数据文件：.txt 且不是本流程的产物"""
    if not name.lower().endswith(".txt"):
        return False
    return not any(name.endswith(s) for s in ("_trimmed.txt", "_fitted_DRT.txt"))


def index_of(name):
    """取文件名中最后一串数字 -> (前缀, 序号, 尾巴)。'... run 10.txt' -> ('... run ', 10, '.txt')"""
    m = re.match(r"^(.*?)(\d+)([^\d]*)$", name)
    if not m:
        return None
    return m.group(1), int(m.group(2)), m.group(3)


def collect_originals(d, explicit):
    """返回 {序号: 原始文件名}，同号取**点数最多**的那个（完整源文件优先）"""
    if explicit:
        dirs = [explicit]
    else:
        dirs = sorted(glob.glob(os.path.join(d, "_backup_*"))) or \
               [os.path.dirname(os.path.abspath(d))]
    found = {}
    for dd in dirs:
        if not os.path.isdir(dd):
            print("[WARN] 原始数据目录不存在，跳过: %s" % dd)
            continue
        for p in sorted(glob.glob(os.path.join(dd, "*.txt"))):
            n = os.path.basename(p)
            if not is_data_file(n):
                continue
            r = index_of(n)
            if not r:
                continue
            # 完整性打分：点数多者优先（数据无关）
                try:
                    score = sum(1 for _ in open(p, encoding="utf-8-sig", errors="ignore"))
                except Exception:
                    score = 0
            k = r[1]
            if k not in found or score > found[k][1]:
                found[k] = (n, score)
    return {k: v[0] for k, v in found.items()}, dirs


def common_prefix(names):
    if not names:
        return ""
    pre = os.path.splitext(names[0])[0]
    for n in names[1:]:
        s = os.path.splitext(n)[0]
        i = 0
        while i < min(len(pre), len(s)) and pre[i] == s[i]:
            i += 1
        pre = pre[:i]
    return pre.rstrip(" _-")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("result_dir")
    ap.add_argument("--stem", default=None, help="体系名，不带开头的 EIS")
    ap.add_argument("--original-dir", default=None)
    ap.add_argument("--apply", action="store_true", help="真正执行（默认只打印计划）")
    a = ap.parse_args()

    d = os.path.abspath(a.result_dir)
    if not os.path.isdir(d):
        print("ERROR: 结果目录不存在: %s" % d)
        return 2

    orig, orig_dirs = collect_originals(d, a.original_dir)
    if not orig:
        print("ERROR: 找不到原始数据文件；请用 --original-dir 指定。")
        return 2
    stem = a.stem or strip_eis(common_prefix(list(orig.values())))
    if not stem:
        print("ERROR: 无法自动确定 STEM，请用 --stem 指定。")
        return 2

    print("结果目录  : %s" % d)
    print("原始数据  : %s" % "; ".join(orig_dirs))
    print("STEM      : %r   （如不对请用 --stem 覆盖）" % stem)
    print("原始文件  : %d 个" % len(orig))
    print("序号映射（内部序号 -> 原始文件名；用于 cycleNN / NN 类旧名，属推断，请核对）:")
    for k in sorted(orig):
        print("   %-4s -> %s" % (k, orig[k]))
    print()

    plan, unknown = [], []
    orig_set = set(orig.values())

    def add(src, dst, why):
        if os.path.abspath(src) != os.path.abspath(dst):
            plan.append((os.path.relpath(src, d), os.path.relpath(dst, d), why))

    for root, dirs, files in os.walk(d):
        rel = os.path.relpath(root, d)
        top = rel.split(os.sep)[0]
        if top.startswith("_backup_"):
            continue                      # 备份：一个字节都不动
        in_canon = rel != "." and top in CANON_DIRS
        for n in sorted(files):
            src = os.path.join(root, n)

            if WORKFILE_RE.match(n):
                add(src, os.path.join(d, "_工作文件", n), "中间件")
                continue
            if n.lower().endswith(".py"):
                add(src, os.path.join(d, "06_脚本", n), "脚本")
                continue
            if n == "DRT_batch_summary.csv":
                add(src, os.path.join(d, "04_汇总表", n), "汇总 csv")
                continue
            if n.startswith("EIS-DRT") and n.lower().endswith((".xlsx", ".csv")):
                ext = os.path.splitext(n)[1]
                add(src, os.path.join(d, "04_汇总表", "EIS-DRT %s%s" % (stem, ext)), "汇总表名归一")
                continue
            if n in FIG_MAP:
                add(src, os.path.join(d, "05_图", FIG_MAP[n].format(stem=stem)), "图名归一")
                continue
            if n.startswith("DRT拟合报告") and n.lower().endswith(".md"):
                add(src, os.path.join(d, "07_报告", "DRT拟合报告_%s.md" % stem), "报告")
                continue

            # 样品级：已经是 <S>_... 且 S 命中原始文件名
            hit = next((s for s in SAMPLE_SUFFIXES
                        if n.endswith(s) and (n[:-len(s)] + ".txt") in orig_set), None)
            if hit:
                sub = {"_trimmed.txt": ("01_去尾后数据", "_trimmed.txt"),
                       "_fitted_DRT.txt": ("02_拟合结果_DRT格式", "_fitted_DRT.txt"),
                       "_residuals.csv": ("03_残差表", "_residuals.csv")}.get(hit, ("05_图", hit))
                add(src, os.path.join(d, sub[0], n[:-len(hit)] + sub[1]), "归位（名字已合规）")
                continue

            # 样品级：cycleNN_* / NN_* 旧名 -> 溯源改名
            m = re.match(r"^(?:cycle|sample)?[ _-]*0*(\d+)(.*)$", os.path.splitext(n)[0], re.I)
            if m and int(m.group(1)) in orig:
                S = os.path.splitext(orig[int(m.group(1))])[0]
                rest = m.group(2).lower()
                if rest.endswith("_drt") or n.endswith("_DRT.txt"):
                    dst = os.path.join(d, "02_拟合结果_DRT格式", S + "_fitted_DRT.txt")
                elif "trim" in rest:
                    dst = os.path.join(d, "01_去尾后数据", S + "_trimmed.txt")
                elif "resid" in rest:
                    dst = os.path.join(d, "03_残差表", S + "_residuals.csv")
                else:
                    unknown.append((os.path.relpath(src, d), "疑似样品级但后缀无法判定"))
                    continue
                add(src, dst, "溯源改名（%d -> %s）" % (int(m.group(1)), os.path.basename(dst)))
                continue

            if not in_canon:
                unknown.append((os.path.relpath(src, d), "未识别（H16：根目录不得有文件）"))

    # ---- 目标重复 / 冲突检查 ----
    by_dst, dup = {}, []
    for src_rel, dst_rel, why in plan:
        if dst_rel in by_dst:
            dup.append((by_dst[dst_rel], src_rel, dst_rel))
        else:
            by_dst[dst_rel] = src_rel
    conflicts = []
    for src_rel, dst_rel, why in plan:
        dst = os.path.join(d, dst_rel)
        src = os.path.join(d, src_rel)
        if os.path.exists(dst) and md5(src) != md5(dst):
            conflicts.append((src_rel, dst_rel))

    print("=" * 108)
    print("计划重排 %d 项：" % len(plan))
    for src_rel, dst_rel, why in plan:
        print("  %-56s -> %-56s [%s]" % (src_rel, dst_rel, why))

    if dup:
        print("\n目标重复 %d 组（多个源映射到同一目标；内容相同则去重，不同则 --apply 时跳过）：" % len(dup))
        for a1, b1, t in dup:
            same = md5(os.path.join(d, a1)) == md5(os.path.join(d, b1))
            print("  %s\n    + %s\n    -> %s   [%s]" % (a1, b1, t, "内容相同，去重" if same else "内容不同，需人工判断"))
    if conflicts:
        print("\n冲突 %d 项（目标已存在且内容不同，--apply 会跳过，绝不覆盖）：" % len(conflicts))
        for s, t in conflicts:
            print("  %s  ->  %s" % (s, t))
    if unknown:
        print("\n需人工判断 %d 项：" % len(unknown))
        for s, why in unknown:
            print("  %-56s [%s]" % (s, why))

    print("\n目标目录结构（§8.0.3 H16）：")
    for c in CANON_DIRS:
        print("   %s/" % c)
    print("   _backup_<DATE>/   （原样保留，本脚本不触碰）")

    if not a.apply:
        print("\n[DRY-RUN] 未改动任何文件。确认计划后加 --apply 执行。")
        return 0

    # ---- 执行 ----
    moved = dedup = skipped = 0
    for c in CANON_DIRS:
        os.makedirs(os.path.join(d, c), exist_ok=True)
    handled_dst = set()
    for src_rel, dst_rel, why in plan:
        src, dst = os.path.join(d, src_rel), os.path.join(d, dst_rel)
        if dst_rel in handled_dst:                       # 重复目标：合并处理
            if os.path.exists(src) and os.path.exists(dst) and md5(src) == md5(dst):
                os.remove(src); dedup += 1
            else:
                print("[SKIP 重复目标且内容不同] %s" % src_rel); skipped += 1
            continue
        if not os.path.exists(src):
            continue
        if os.path.exists(dst):
            if md5(src) == md5(dst):
                os.remove(src); dedup += 1
            else:
                print("[SKIP 冲突] %s" % src_rel); skipped += 1
            continue
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.move(src, dst)
        handled_dst.add(dst_rel)
        moved += 1
    print("\n[APPLIED] 移动/改名 %d 项，去重删除 %d 项，跳过 %d 项。" % (moved, dedup, skipped))
    print("提醒：重命名后请重跑 `DRT汇总表_H14校验.py`，确认没有整列被 SKIP。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
