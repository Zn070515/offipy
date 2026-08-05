# 审计固定资产说明（tests/fixtures/audit/）

固定验收集，一次性 git 入库。改动审计逻辑后需手动重跑生成脚本并复核，CI
**不**要求每次字节级重建。

| 文件 | 来源 | 用途 |
|------|------|------|
| `synthetic.pptx` | `generate_audit_fixtures.py` | 一条规则一个场景：越界 HIGH / off-canvas / 贴边 / 部分重叠 / 文本被覆盖 / nowrap / 显式多行 / 页码 / 全页背景 / normAutofit / spAutoFit |
| `edge_cases.pptx` | `generate_audit_fixtures.py` | 误报控制：缩放 group / 嵌套 group / rotation 45° / flipH group / flipV / hidden / connector / table |
| `baseline.pptx` | `generate_audit_fixtures.py` | 基线回归对比对象 |
| `candidate.pptx` | `generate_audit_fixtures.py` | 对比候选：新增越界 + 修复贴边 + 移动/缩放 + 文本变化 + 增删 shape |
| `deck_generated.pptx` | `offipy deck` 渲染示例 starter（需 chromium） | 真实管线产物审计（不做过度强断言） |

## 再生成

```bash
uv run python tests/fixtures/audit/generate_audit_fixtures.py
```

`deck_generated.pptx` 不在此脚本内（需 chromium），用 `offipy deck render`
对示例 HTML 渲染一次后拷入；内容与审计结论仅供真实验收参考。

## 资产不变量

`tests/test_audit_fixtures.py` 对这些文件做**固定断言**，改动时注意保持一致：

- synthetic 的关键 rule_id 与严重度、suppressed 豁免原因、零 warning；
- edge_cases **零 finding/suppressed/warning**（connector/hidden/rotate/flip/group 不误判），
  且 Affine2D 组几何精确（缩放 group 子局部坐标 → 幻灯片绝对坐标）；
- baseline→candidate 回归：新增 `bounds.partial` MID、已解决 `margin.right`、
  形状增删/移动/缩放/文本变化、门禁严重度 MID。

## 生成脚本踩坑（勿回退）

python-pptx **往嵌套 group 里加子 shape 会重置外层 group 的 `p:grpSpPr/a:xfrm`
为 off=(0,0) ext=(1,1)**。因此外层 group 的 left/top/width/height 必须在所有子
（含嵌套 group）都加完之后再设置，`a:chExt` 最后强制覆盖——否则 group 几何错位，
edge_cases 的 Affine2D 精确断言会失败。
