面向 **DeepSeek Harness** 的 DRT（弛豫时间分布）拟合全流程硬规范：给出 EIS 阻抗谱数据（f / Z' / Z''）即可产出可复现、可发表的 DRT 结果。

## 内容

- **§1–§8 全流程**：输入体检与去尾（口径 A/B）、模型与 τ 网格（含 L 判定）、λ 选择、残差与一致性校验、出图、汇总表、固定命名与目录归档
- **H1–H21 硬门槛**：每条都写明判据、失败处置与「不许做什么」，含 H1′ / H5′ / H7′ / H8′ 配套条款
- **附录 A**：拟合核心参考实现（Python，可直接执行）
- **附录 B**：方法学依据（不含任何原始数据）
- **附录 D**：归档工具箱的逐字节副本
- **归档工具箱**：`summary` / `check` / `normalize` / `numbers` / `verify` 五个子命令
- **自检脚本**：协议包一致性（含根目录结构守卫）+ 工具箱回归（全合成数据）

## v1.20 修正

本版由一次系统性文本自查驱动，修正 6 处问题：

| 项 | 内容 |
|---|---|
| 许可措辞 | 资源块由「专有」改为「Apache-2.0 开源」，与 `LICENSE` / `NOTICE` 一致 |
| 工具箱版本串 | `PROTOCOL` 常量内部版本由 v1.17 修正为 v1.20，并移除旧的「LICENSE（专有）」表述（该串会写入汇总表 `Info` 页） |
| **自检增强** | `P2` 现在断言 `PROTOCOL` 串内版本 == `PROTOCOL_VERSION` —— 此前只提取不比较，属校验盲区；已用「改错必 FAIL」负向验证 |
| 条款同步 | §6 诊断表「切弧」行由 κ 单判据更新为**切弧三判据**（κ / 宽窄窗峰数对比 / 窄窗 `γ_max` 距慢端） |
| 交叉引用 | 附录 B 的去尾证据规则由「H1 第 4 条」修正为「§1 第 5 条」 |
| 消歧 | §4.2 补充：组内跨样品仍须固定同一 λ（H1.4），与「λ 不可跨批次照搬」不冲突 |

自查范围与结论：24 个 `§x.y` 引用全部可解析；H1–H21 全部在 §5 有定义；10 张表列数一致；排版（行尾空白 / 制表符 / 连续标点 / NUL）0 问题；陈旧词 0；**数据痕迹 0**。

## 安装

把 `drt-fitting-protocol/` 整个目录放进 DSH 扫描的 skill 目录（项目级 `.dsh/skills/` 或用户级 `<dshHome>/skills/`，以 DSH 官方文档为准）；也可以直接解压下方 Assets 里的 zip。

## 运行环境

DSH 自带 Python 已包含全部依赖（numpy / scipy / matplotlib / openpyxl / pandas），**无需安装任何包**：先调用 `load_workspace_dependencies` 取得 Python 路径，再执行 `scripts/` 下的脚本。

## 校验

```bash
python scripts/自检_协议包.py        # 协议包结构与一致性
python scripts/自检_工具箱回归.py     # 工具箱回归（合成数据）
```

## 关于本附件

附件 `drt-fitting-protocol.zip` 由仓库内的 GitHub Actions 在推 tag 时**自动构建并上传**，构建脚本为 `tools/build_zip.py`：

- **可复现**：固定时间戳与条目顺序，同样输入必定得到同样字节（CI 每次跑两遍比对 sha256）
- **发布前已过门禁**：换行检查 → 可复现性 → 包自检（结构 / 版本串 / 附录 D 逐字节 / 引用完整性 / 编码）→ 工具箱回归 → 隐私体检
- 附件与仓库内 `drt-fitting-protocol/` 目录逐文件一致

## 许可

Apache License 2.0 —— 见 [LICENSE](https://github.com/yangminghu/drt-fitting-protocol/blob/main/LICENSE) 与 [NOTICE](https://github.com/yangminghu/drt-fitting-protocol/blob/main/NOTICE)。

本仓库与发布附件均**不含任何实验数据**；回归测试全部使用合成数据。
