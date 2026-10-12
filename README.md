# DRT 拟合规范（drt-fitting-protocol）

> 一套**硬门槛化**的 DRT（弛豫时间分布）拟合全流程规范：给出 EIS 阻抗谱数据（f / Z' / Z''）即可产出**可复现、可发表**的 DRT 结果。

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)

- **版本**：v1.20
- **许可**：[Apache License 2.0](LICENSE)（开源；允许商用与闭源衍生，须保留版权与许可声明）
- **触发词**：DRT、弛豫时间分布、DRT 拟合、DRT 分析、EIS 去尾、阻抗谱拟合、R_pol、拟合结果归档

## 这是什么

`SKILL.md` 定义了一套分节（§1–§8）的拟合流程与 **H1–H21 硬门槛**：去尾（口径 A/B）、模型与 τ 网格、λ 选择、残差与一致性校验、出图、汇总表、固定命名与目录归档。每条门槛都写明判据、失败处置与「不许做什么」，用于避免 DRT 分析中最常见的几类错误 —— 低频扩散尾未去除、跨样品统一窗口把弧切掉、夹具电感污染被当成样品特征、残差超限仍照常交付、命名与归档不规范等。

## 目录结构

```
drt-fitting-protocol/
├── SKILL.md                      # 技能定义（frontmatter + 协议正文 §0–§8 + 附录 A 参考实现）
├── references/                   # 参考资料（按需读取）
│   ├── 附录B_实证依据.md
│   ├── 附录D_归档工具箱备份.md
│   └── 许可与署名.md
└── scripts/                      # 可执行脚本
    ├── DRT归档工具箱.py           # 唯一权威实现（summary / check / normalize / numbers / verify）
    ├── 自检_协议包.py
    ├── 自检_工具箱回归.py
    └── 同步附录D_维护脚本.py 等维护脚本
```

## 使用方法

1. 把 `drt-fitting-protocol/` 整个目录放入技能目录（例如 Agent Skills 的 `skills/` 下，或打包上传到 WorkBuddy 开放平台）。
2. 让 AI 加载 `SKILL.md`，直接提供数据（`f / Z' / Z''` 三列文本），按 §1–§8 执行。
3. 归档工具箱（Bash）：

```bash
python scripts/DRT归档工具箱.py summary "<结果根>/02_拟合结果_DRT格式" "EIS-DRT <STEM>" \
    --out-dir "<结果根>/04_汇总表" --note "<一句话>"
python scripts/DRT归档工具箱.py check   "<结果根>/04_汇总表/EIS-DRT <STEM>.xlsx" --dir "<结果根>/02_拟合结果_DRT格式"
python scripts/DRT归档工具箱.py numbers "<结果根>"
python scripts/DRT归档工具箱.py verify  "<结果根>"
```

## 运行环境

- Python 3.9+
- `numpy`、`openpyxl`（归档工具箱）；`scipy`（附录 A 拟合核心）；`matplotlib`（出图）
- 脚本均为 `python xxx.py` 形式，不含 Windows 专用写法，在 Bash 与 PowerShell 下都可执行

### 在 DeepSeek Harness 中运行（推荐）

本 skill 面向 **DeepSeek Harness** 制作。DSH 自带的 Python 运行时**已包含上述全部依赖**
（实测 3.12：numpy / scipy / matplotlib / openpyxl / pandas），因此**无需安装任何包**：
先调用 `load_workspace_dependencies` 取得自带 Python 的绝对路径，再用它执行 `scripts/` 下的脚本。

安装位置：把 `drt-fitting-protocol/` 整个目录放进 DSH 会扫描的 skill 目录 —— 项目级的
`.dsh/skills/`，或用户级的 `<dshHome>/skills/`（具体扫描范围与优先级**以 DSH 官方文档为准**）。

> 技能目录应视为**只读资源**；中间产物与交付物请写到任务工作区。

### 在其他环境中运行

`SKILL.md` 是跨 agent 的通用格式，本包同样可用于 Claude Code、Codex、WorkBuddy 等支持
`SKILL.md` 的环境。WorkBuddy 开放平台另有上传规范（根目录仅 `SKILL.md` + `references/` +
`scripts/` + `templates/`），本包的结构符合该规范。

## 自检

```bash
python scripts/自检_协议包.py        # 协议包结构与一致性（含根目录结构守卫）
python scripts/自检_工具箱回归.py     # 工具箱回归测试（合成数据）
```

## 自动校验与发布

仓库带 GitHub Actions（[`.github/workflows/ci.yml`](.github/workflows/ci.yml)）：

- **推 `main` / 提 PR**：换行检查 → 确定性打包（跑两遍比对 sha256）→ 包自检 → 工具箱回归 → 隐私体检 → 上传产物
- **推 `v*` tag**：在上面全部通过后，自动**创建或更新 Release**，并把 `drt-fitting-protocol.zip` 挂成附件（同时清掉改名后残留的过期附件）

打包用 [`tools/build_zip.py`](tools/build_zip.py)：固定时间戳与条目顺序，**同样输入必定得到同样字节** —— 所以本地打的包与 Release 附件逐字节相同，可复现。

发布新版本：

```bash
# 1) 改包内容，并把版本串升到位（SKILL.md / 工具箱 / 附录 D 的一致性由自检脚本把关）
# 2) 提交
git add -A && git commit -m "..." && git push
# 3) 打 tag 推送 —— 剩下的交给 CI
git tag -a v1.21 -m "DRT 拟合规范 v1.21" && git push origin v1.21
```

> 隐私体检的具体禁词**不写在仓库里**（那等于把要藏的东西公开），通过仓库变量 `PRIVACY_DENY`（Settings → Secrets and variables → Actions → Variables）传入；留空则只跑通用规则：数据类扩展名 / 绝对路径 / 家目录 / 邮箱 / 手机号 / NUL。

## 仓库说明

- 本仓库**不包含任何实验数据**；协议中的示例与回归测试均使用**合成数据**。
- `.gitignore` 已排除 `*.txt`、`*.zip`、`*.xlsx`、`*.csv` 等，避免误提交实测数据。
- 本仓库采用 **Apache License 2.0**（开源）：允许商用与闭源衍生，须保留版权与许可声明、标注修改、并随附 [NOTICE](NOTICE)。
- 本仓库与 Anthropic、DeepSeek 及其他任何组织**无隶属、赞助或背书关系**；文中提及的产品名称仅用于说明运行环境与兼容性，相关商标归各自所有者。
- 文档中关于运行环境的描述为作者自述，如与官方文档不一致，**以官方文档为准**。

## 许可

[Apache License 2.0](LICENSE)

- **可以**：自由使用、复制、修改、合并、发布、分发、再许可（含商用、含闭源衍生作品）。
- **须遵守**：随附许可证副本；修改过的文件显著标注已修改；保留版权/专利/商标/归属声明；随附 [NOTICE](NOTICE)；不得以「Apache」或著作权人名义为衍生作品背书。
- **免责**：按「现状」提供，不附带任何明示或默示担保。

```
Copyright 2026 Yangming Hu

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
