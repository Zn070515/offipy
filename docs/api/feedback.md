> [English](feedback.en.md)

# 反馈学习 API

### `train`

离线训练反馈学习系统：读 ~/.offipy/art_feedback.jsonl → 按当前 FEATURES schema 编码 → 构配对（同 rule×profile：fixed>accepted）→ numpy MLP 训练 → 原子写 art_feedback_model.json。样本不足/无有效样本时返回状态而非报错（不删除已有模型）。需要 numpy：pip install "offipy[feedback]"。

- **参数**: `feedback_dir: str`、`seed: int`
- **返回**: `dict`
- **标志**: 普通操作

---

### `status`

反馈学习状态：样本数、可配对潜力、当前模型状态（none/valid/expired/stale/corrupt）。只读本机数据，不训练不写文件。

- **参数**: `feedback_dir: str`
- **返回**: `dict`
- **标志**: 只读

---

### `append`

追加一条反馈标签：用户对某 rule 的 finding 处置（fixed=该修，accepted=规则判断对，ignored=无关）。写入 feedback_dir 的 JSONL（未给 feedback_dir 时写默认 ~/.offipy），供 feedback train 学习。features 为扁平特征快照（encode_features 产出），CLI 传 JSON 字符串。

- **参数**: `profile: str`、`rule_id: str`、`action: str`、`severity: str`、`slide_index: int`、`message: str`、`source: str`、`feedback_dir: str`、`ts: str`、`features: any`、`feature_schema_version: str`
- **返回**: `dict`
- **标志**: 普通操作

---

### `recommend`

只读建议：对 .pptx 跑艺术分析 + 学习推理，返回调整 finding 与确定性建议（不进文档、不写反馈库）。需要有效模型：无模型/过期/损坏时显式报错（不回退 v2）。可用 --json（generic dispatch 恒输出 JSON）。

- **参数**: `pptx: str`、`feedback_dir: str`、`profile: str`、`json: bool`
- **返回**: `dict`
- **标志**: 只读

---

### `apply`

把学习到的 rule.delta 持久化到 profile 存储（默认 ~/.offipy/art_profiles.json），使 `deck audit --profile <name>`（不带 --feedback-dir）也反映学习调整。需要有效模型。

- **参数**: `profile: str`、`feedback_dir: str`
- **返回**: `dict`
- **标志**: 普通操作

---

### `reschema`

把 schema 过期的历史反馈记录（有特征快照）原地重写为当前 feature_schema_version，供 bump 后迁移。返回 {rewritten, skipped_no_features, already_current}。坏行保留不破坏文件。

- **参数**: `feedback_dir: str`
- **返回**: `dict`
- **标志**: 普通操作
