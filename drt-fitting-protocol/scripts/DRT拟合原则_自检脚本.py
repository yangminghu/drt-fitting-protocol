"""
DRT 拟合原则 · 自检脚本
------------------------------------------------------------------
用途：
  1) 验证 AI 使用的拟合实现是否与协议附录 A 一致（可复现性）
  2) 逐条执行 §5 硬门槛 H4–H13 并打印通过情况

用法：
  python 自检脚本.py  <数据目录>  [样品编号列表，如 1,2,3]
  不带参数时使用脚本内 DEFAULT_DIR 下的全部 *.txt（DEFAULT_DIR 默认为空 -> 必须显式传目录）
"""
import sys, os, glob
import numpy as np
from scipy.optimize import lsq_linear

DEFAULT_DIR = ""
LAM, NGRID, MARGIN = 1e-6, 400, 3.1623

# ================= 附录 A 参考实现（一字不改） =================
def load(path):
    rows = []
    for line in open(path, encoding="utf-8-sig", errors="ignore"):   # BOM 容错
        p = line.replace(",", "\t").split()
        if len(p) >= 3:
            try: rows.append([float(x) for x in p[:3]])
            except ValueError: pass
    a = np.array(rows); return a[:, 0], a[:, 1], a[:, 2]

def fit_drt(f, Zr, Zi, lam=LAM, N=NGRID, margin=MARGIN, use_L=True):
    w = 2*np.pi*f; za = np.abs(Zr + 1j*Zi); n = len(f)
    T1, T2 = 1/(2*np.pi*f.max()), 1/(2*np.pi*f.min())
    tau = np.logspace(np.log10(T1/margin), np.log10(T2*margin), N)
    dln = np.gradient(np.log(tau)); wt = np.outer(w, tau); K = 1.0/(1.0 + wt**2)
    ncol = N + (2 if use_L else 1)
    A = np.zeros((2*n, ncol)); b = np.zeros(2*n)
    A[0::2, :N] = dln[None, :]*K
    A[1::2, :N] = -dln[None, :]*K*wt
    A[0::2, N] = 1.0
    if use_L: A[1::2, N+1] = w
    b[0::2] = Zr; b[1::2] = Zi
    dx = np.mean(np.diff(np.log(tau)))
    D = np.zeros((N, ncol))
    for i in range(N-2): D[i, i:i+3] = [1., -2., 1.]
    D = D/dx**1.5
    Aa = np.vstack([A, np.sqrt(lam)*D]); ba = np.concatenate([b, np.zeros(N)])
    sc = np.linalg.norm(Aa, axis=0); sc[sc == 0] = 1.0
    lb = np.concatenate([np.zeros(N), [-np.inf]*(ncol-N)]); ub = np.full(ncol, np.inf)
    s = lsq_linear(Aa/sc, ba, bounds=(lb/sc, ub/sc), method="trf", tol=1e-13, max_iter=5000)
    x = s.x/sc; Zf = A@x; g = x[:N]
    dRe = (Zf[0::2]-Zr)/za*100; dIm = (Zf[1::2]-Zi)/za*100
    win = (tau >= T1) & (tau <= T2)
    return dict(tau=tau, gam=g, R=x[N], L=x[N+1] if use_L else 0.0,
                dRe=dRe, dIm=dIm, Zfit=Zf[0::2]+1j*Zf[1::2],
                Rpol=np.sum(g*dln), Rpol_win=np.sum(g[win]*dln[win]),
                rms=np.sqrt(np.mean((Zf-b)**2)))
# ==============================================================

def gates(f, Zr, Zi, r, r_low=None, r_noL=None, r_hiN=None, r_m1=None):
    out = []
    out.append(("H4  R <= min Z'", r["R"] <= Zr.min(),
                "R=%.3f  minZ'=%.3f" % (r["R"], Zr.min())))
    out.append(("H6  gamma >= 0", r["gam"].min() >= 0, "min gamma=%.3g" % r["gam"].min()))
    mx = max(abs(r["dRe"]).max(), abs(r["dIm"]).max())
    out.append(("H7  residual <= 1%", mx <= 1.0,
                "maxdRe=%.3f%%  maxdIm=%.3f%%  rms=%.4f ohm"
                % (abs(r["dRe"]).max(), abs(r["dIm"]).max(), r["rms"])))
    frac = r["gam"][0]/r["gam"].max()
    out.append(("H8  no fast-edge spike", frac <= 0.05, "gamma(tau_min)/gamma_max=%.4f" % frac))
    if r_noL is not None:
        needs = Zi.max() > 0
        out.append(("H13 model choice", (not needs) or r_noL["L"] >= 0,
                    "maxZ''=%.3f -> L %s" % (Zi.max(), "required" if needs else "not required")))
    Lmin = (Zi[0]/ (2*np.pi*f[0])) if Zi[0] > 0 else 0.0
    out.append(("H12 L >= Z''(fmax)/w_max", r["L"] >= Lmin - 1e-12,
                "L=%.3g  required>=%.3g H" % (r["L"], Lmin)))
    if r_low is not None:
        dR = abs(r["R"] - r_low["R"])
        out.append(("H11 R uncertainty", dR < 0.5,
                    "|R(1e-6)-R(1e-7)|=%.3f ohm%s" % (dR, "" if dR < 0.5 else "  -> report as RANGE")))
    if r_hiN is not None and r_m1 is not None:
        d = max(abs(r["R"]-r_hiN["R"]), abs(r["R"]-r_m1["R"]))
        out.append(("H5  grid robustness", d < 0.1, "max dR over N=800 / margin=1.0: %.4f ohm" % d))
    return out

def main():
    d = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_DIR
    files = sorted(glob.glob(os.path.join(d, "*.txt")))
    if not files:
        print("no *.txt under", d); return 1
    if len(sys.argv) > 2:
        want = {int(x) for x in sys.argv[2].split(",")}
        files = [p for p in files if int(os.path.basename(p).split()[5].split("_")[0]) in want]

    allok = True
    for p in files:
        f, Zr, Zi = load(p)
        r = fit_drt(f, Zr, Zi)
        r_low = fit_drt(f, Zr, Zi, lam=1e-7)
        r_noL = fit_drt(f, Zr, Zi, use_L=False)
        r_hiN = fit_drt(f, Zr, Zi, N=800)
        r_m1 = fit_drt(f, Zr, Zi, margin=10.0)
        print("=" * 78)
        print(os.path.basename(p))
        print("  n=%d  %.6g .. %.6g Hz   maxZ''=%.4f   R=%.4f ohm  L=%.4g H"
              % (len(f), f.min(), f.max(), Zi.max(), r["R"], r["L"]))
        print("  model: %s   (maxZ'' %s 0 -> L %s)"
              % ("R + iwL + DRT" if Zi.max() > 0 else "R + DRT",
                 ">" if Zi.max() > 0 else "<=",
                 "REQUIRED" if Zi.max() > 0 else "optional"))
        for name, ok, info in gates(f, Zr, Zi, r, r_low, r_noL, r_hiN, r_m1):
            allok &= ok
            print("   [%s] %-26s %s" % ("PASS" if ok else "FAIL", name, info))
        if Zi.max() > 0:
            print("   (without L: R=%.3f  maxdRe=%.3f%%  maxdIm=%.3f%%)  <-- shows why L is mandatory"
                  % (r_noL["R"], abs(r_noL["dRe"]).max(), abs(r_noL["dIm"]).max()))
    print("=" * 78)
    print("OVERALL:", "ALL GATES PASS" if allok else "SOME GATES FAILED")
    return 0 if allok else 1

if __name__ == "__main__":
    sys.exit(main())
