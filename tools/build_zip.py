#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""确定性打包：同样的输入必定得到同样的字节（可复现构建）。

固定项：条目顺序、时间戳、压缩方式与级别、外部属性、create_system、无目录条目。
因此「本地打出来的 zip」与「CI 打出来的 zip」逐字节相同，Release 附件可复现。

用法：
    python tools/build_zip.py                                  # 默认包目录 -> dist/drt-fitting-protocol.zip
    python tools/build_zip.py --pkg <包目录> --out <zip 路径>
退出码：0 成功；1 包结构与规范清单不一致。
"""
import argparse
import os
import sys
import zipfile

ROOT = "drt-fitting-protocol"

# 包内文件清单（= WorkBuddy 技能规范结构：SKILL.md + references/ + scripts/）
ORDER = [
    "SKILL.md",
    "references/附录B_实证依据.md",
    "references/附录D_归档工具箱备份.md",
    "references/许可与署名.md",
    "scripts/DRT归档工具箱.py",
    "scripts/DRT拟合原则_自检脚本.py",
    "scripts/DRT汇总表_H14校验.py",
    "scripts/DRT汇总表_生成脚本.py",
    "scripts/DRT输出命名规范化.py",
    "scripts/同步主文档_维护脚本.py",
    "scripts/同步附录D_维护脚本.py",
    "scripts/自检_协议包.py",
    "scripts/自检_工具箱回归.py",
]

FIXED_DT = (2026, 1, 1, 0, 0, 0)      # 固定时间戳 —— 可复现的关键
EXTERNAL_ATTR = 0o100644 << 16        # -rw-r--r--
CREATE_SYSTEM = 0                     # 与原包一致


def inventory(pkg):
    """返回 (缺失清单, 多余清单)，用于拦住"包被改脏了还在打包"。"""
    have = set()
    for dp, dn, fn in os.walk(pkg):
        for f in fn:
            rel = os.path.relpath(os.path.join(dp, f), pkg).replace("\\", "/")
            have.add(rel)
    want = set(ORDER)
    return sorted(want - have), sorted(have - want)


def build(pkg, out):
    missing, extra = inventory(pkg)
    if missing or extra:
        print("  [FAIL] 包结构与规范清单不一致")
        if missing:
            print("         缺失:", missing)
        if extra:
            print("         多余:", extra)
        return 1
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for rel in ORDER:
            src = os.path.join(pkg, rel.replace("/", os.sep))
            zi = zipfile.ZipInfo(ROOT + "/" + rel, date_time=FIXED_DT)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = EXTERNAL_ATTR
            zi.create_system = CREATE_SYSTEM
            with open(src, "rb") as fh:
                z.writestr(zi, fh.read())
    n = len(ORDER)
    size = os.path.getsize(out)
    print("  [OK] %s -> %s（%d 条目 / %d 字节）" % (pkg, out, n, size))
    return 0


def main():
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ap = argparse.ArgumentParser(description="确定性打包技能包")
    ap.add_argument("--pkg", default=os.path.join(here, ROOT), help="技能包目录")
    ap.add_argument("--out", default=os.path.join(here, "dist", ROOT + ".zip"), help="输出 zip 路径")
    a = ap.parse_args()
    return build(a.pkg, a.out)


if __name__ == "__main__":
    sys.exit(main())
