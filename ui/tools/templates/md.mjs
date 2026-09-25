// Markdown templates for the new-file gallery: build returns the file's text, in the language t picks.
export const cats = [['笔记', 'Notes'], ['工作', 'Work'], ['技术', 'Tech']];

export default [
  {
    id: 'notes', cat: '笔记', name: ['学习笔记', 'Study Notes'],
    build: (b, t) => t(`
# 学习笔记：用户研究入门

> 来源：线上课程第 3 讲　　日期：2026-09-25　　标签：#产品 #研究

## 本讲要点

1. **先定问题，再选方法**：研究问题决定用访谈、问卷还是可用性测试。
2. **访谈听行为，不听观点**：问「上一次是怎么做的」，少问「你会不会」。
3. **样本不在多**：同一类用户访谈 5–8 人，就能看到大部分问题。

## 方法对比

| 方法 | 适合回答 | 样本量 | 周期 |
| --- | --- | --- | --- |
| 深度访谈 | 为什么、怎么做 | 5–8 人 | 1–2 周 |
| 问卷调查 | 有多少、比例如何 | 200 人以上 | 1 周 |
| 可用性测试 | 哪里卡住了 | 5 人 | 3–5 天 |

## 我的理解

访谈像是在收集「故事」，问卷是在数「故事出现了多少次」。先有故事，才知道问卷该问什么。

## 待办

- [x] 看完第 3 讲
- [ ] 整理一份访谈提纲
- [ ] 找 3 位同事试访

## 疑问

- 线上访谈和面对面访谈，信息量差多少？
`, `
# Study Notes: An Introduction to User Research

> Source: online course, lesson 3　Date: 2026-09-25　Tags: #product #research

## Key points

1. **Question first, method second**: the research question decides between interviews, surveys and usability tests.
2. **Listen for behaviour, not opinions**: ask "how did you do it last time?", not "would you use it?".
3. **Small samples go far**: five to eight interviews per user group surface most problems.

## Methods compared

| Method | Answers | Sample | Time |
| --- | --- | --- | --- |
| In-depth interviews | Why and how | 5–8 people | 1–2 weeks |
| Survey | How many, what share | 200+ | 1 week |
| Usability test | Where people get stuck | 5 people | 3–5 days |

## My take

Interviews collect the stories; surveys count how often each story happens. You need the stories before you know what to ask in a survey.

## To do

- [x] Watch lesson 3
- [ ] Draft an interview guide
- [ ] Run practice interviews with three colleagues

## Questions

- How much is lost when an interview happens over video instead of in person?
`),
  },
  {
    id: 'weekly-plan', cat: '工作', name: ['周计划', 'Weekly Plan'],
    build: (b, t) => t(`
# 周计划 · 2026 年第 40 周（9 月 28 日 – 10 月 2 日）

## 本周三件大事

1. [ ] 新版官网信息架构定稿并评审通过
2. [ ] 秋季新品首发活动物料全部到位
3. [ ] 完成三季度复盘报告初稿

## 每日安排

| | 上午 | 下午 |
| --- | --- | --- |
| 周一 | 周会；整理上周待办 | 官网信息架构修订 |
| 周二 | 用户访谈 × 2 | 访谈记录整理 |
| 周三 | 信息架构评审会 | 首发活动物料核对 |
| 周四 | 复盘报告：数据部分 | 与渠道伙伴电话会 |
| 周五 | 复盘报告：结论与建议 | 周总结；下周计划 |

## 待办清单

### 必须完成
- [ ] 评审会材料（周二前发出）
- [ ] 物料到货清单核对
- [ ] 复盘报告初稿

### 尽量完成
- [ ] 更新渠道伙伴培训包
- [ ] 回复 3 条客户反馈

### 可以延后
- [ ] 整理设计规范文档

## 周末回顾（周五填写）

- 完成了什么？
- 什么没按计划走，为什么？
- 下周第一件事是什么？
`, `
# Weekly Plan · Week 40, 2026 (September 28 – October 2)

## Three things that matter this week

1. [ ] Final information architecture for the new website, approved in review
2. [ ] Every piece of launch material for the autumn campaign in place
3. [ ] First draft of the Q3 review

## Day by day

| | Morning | Afternoon |
| --- | --- | --- |
| Monday | Team meeting; clear last week's to-dos | Revise the information architecture |
| Tuesday | User interviews × 2 | Write up the interviews |
| Wednesday | Architecture review | Check the launch materials |
| Thursday | Q3 review: the numbers | Call with the channel partners |
| Friday | Q3 review: conclusions | Weekly wrap-up; plan next week |

## To do

### Must
- [ ] Review materials (sent by Tuesday)
- [ ] Reconcile the delivery list
- [ ] Q3 review draft

### Should
- [ ] Update the partner training kit
- [ ] Answer three customer comments

### Could
- [ ] Tidy the design guidelines

## Friday review

- What got done?
- What slipped, and why?
- What is the first thing next week?
`),
  },
  {
    id: 'tech-design', cat: '技术', name: ['技术方案', 'Design Doc'],
    build: (b, t) => t(`
# 技术方案：文档实时协同编辑

| 状态 | 评审中 | 作者 | 刘洋 | 更新 | 2026-09-25 |
| --- | --- | --- | --- | --- | --- |

## 背景

当前多人编辑同一文档时以「最后保存者为准」，冲突每周约 30 起，用户需要手动合并。目标是让多人同时编辑同一份文档时改动实时可见、自动合并、可追溯。

## 目标与非目标

**目标**
- 同一文档 100 人同时编辑，操作延迟 P95 < 300 ms
- 断网期间可继续编辑，恢复后自动合并
- 每次改动可归属到人，可回滚

**非目标**
- 不做跨文档的事务
- 首期不支持评论与批注的实时同步

## 方案概述

采用 CRDT（Yjs）在客户端合并操作，服务端只负责广播与持久化快照：

\`\`\`
客户端 A ──┐                     ┌── 客户端 B
           ├─ WebSocket ─ 同步服务 ─┤
客户端 C ──┘        │              └── 客户端 D
                    ▼
              快照存储（每 5 分钟 / 每 500 次操作）
\`\`\`

- 客户端维护完整文档状态，离线操作进入本地队列
- 同步服务无状态，可水平扩展；房间按文档 ID 分片
- 快照 + 操作日志组合恢复，日志保留 30 天

## 接口设计

| 接口 | 方法 | 说明 |
| --- | --- | --- |
| \`/ws/doc/{id}\` | WebSocket | 加入房间，双向同步更新 |
| \`/api/doc/{id}/snapshot\` | GET | 最新快照，附版本号 |
| \`/api/doc/{id}/history?from=\` | GET | 指定版本后的操作日志 |
| \`/api/doc/{id}/restore\` | POST | 回滚到指定版本 |

## 数据模型

\`\`\`json
{
  "docId": "d_8f3a",
  "version": 1284,
  "snapshot": "<binary>",
  "ops": [{ "seq": 1285, "user": "u_12", "at": "2026-09-25T08:12:03Z", "delta": "<binary>" }]
}
\`\`\`

## 风险与应对

| 风险 | 影响 | 应对 |
| --- | --- | --- |
| 大文档首次加载慢 | 体验差 | 分块加载 + 增量渲染 |
| 同步服务单点 | 全员不可用 | 双活 + 客户端自动重连 |
| 历史膨胀 | 存储成本 | 定期压缩快照，日志 30 天过期 |

## 里程碑

| 阶段 | 内容 | 时间 |
| --- | --- | --- |
| M1 | 原型：两端同步，无持久化 | 10 月 15 日 |
| M2 | 持久化、离线队列、权限接入 | 11 月 10 日 |
| M3 | 灰度 5% 用户，观察一周 | 11 月 24 日 |
| M4 | 全量发布 | 12 月 8 日 |

## 待讨论

- 是否在 M2 前引入操作压缩？
- 匿名协作者的归属如何展示？
`, `
# Design Doc: Real-Time Collaborative Editing

| Status | In review | Author | Casey Kim | Updated | 2026-09-25 |
| --- | --- | --- | --- | --- | --- |

## Background

Today the last save wins when several people edit one document. That produces about 30 conflicts a week, each merged by hand. The goal: everyone editing the same document sees changes live, merged automatically and attributable.

## Goals and non-goals

**Goals**
- 100 people editing one document with P95 operation latency under 300 ms
- Editing continues offline and merges on reconnect
- Every change is attributed and can be rolled back

**Non-goals**
- No transactions across documents
- Comments are not synced live in the first release

## Design

Operations merge on the client with a CRDT (Yjs); the server only broadcasts and stores snapshots:

\`\`\`
client A ──┐                       ┌── client B
           ├─ WebSocket ─ sync service ─┤
client C ──┘          │                └── client D
                      ▼
            snapshot store (every 5 min / 500 ops)
\`\`\`

- The client holds the whole document; offline edits queue locally
- The sync service is stateless and scales out; rooms are sharded by document id
- Recovery replays the log over the last snapshot; logs are kept for 30 days

## API

| Endpoint | Method | Purpose |
| --- | --- | --- |
| \`/ws/doc/{id}\` | WebSocket | Join the room; two-way updates |
| \`/api/doc/{id}/snapshot\` | GET | Latest snapshot with its version |
| \`/api/doc/{id}/history?from=\` | GET | Operations after a version |
| \`/api/doc/{id}/restore\` | POST | Roll back to a version |

## Data model

\`\`\`json
{
  "docId": "d_8f3a",
  "version": 1284,
  "snapshot": "<binary>",
  "ops": [{ "seq": 1285, "user": "u_12", "at": "2026-09-25T08:12:03Z", "delta": "<binary>" }]
}
\`\`\`

## Risks

| Risk | Impact | Mitigation |
| --- | --- | --- |
| Large documents load slowly at first | Poor experience | Chunked loading and incremental rendering |
| The sync service is a single point of failure | Nobody can edit | Two active instances; clients reconnect on their own |
| History grows without bound | Storage cost | Periodic snapshot compaction; logs expire after 30 days |

## Milestones

| Phase | Scope | Date |
| --- | --- | --- |
| M1 | Prototype: two clients syncing, nothing stored | October 15 |
| M2 | Persistence, offline queue, permissions | November 10 |
| M3 | 5% of users for a week | November 24 |
| M4 | Everyone | December 8 |

## Open questions

- Compress operations before M2?
- How do we show attribution for anonymous collaborators?
`),
  },
  {
    id: 'readme', cat: '技术', name: ['项目 README', 'Project README'],
    build: (b, t) => t(`
# 项目名称

一句话说明这个项目做什么、给谁用。

## 功能

- 功能一：做什么，解决什么问题
- 功能二
- 功能三

## 快速开始

\`\`\`bash
git clone https://example.com/org/project.git
cd project
npm install
npm run dev
\`\`\`

打开 http://localhost:3000 即可看到运行结果。

## 配置

| 变量 | 说明 | 默认值 |
| --- | --- | --- |
| \`PORT\` | 服务端口 | \`3000\` |
| \`DATABASE_URL\` | 数据库连接串 | 无 |
| \`LOG_LEVEL\` | 日志级别 | \`info\` |

## 目录结构

\`\`\`
src/        源代码
tests/      测试
docs/       文档
scripts/    构建与发布脚本
\`\`\`

## 参与贡献

1. Fork 本仓库并新建分支
2. 提交前运行 \`npm test\`
3. 发起 Pull Request，说明改动与原因

## 许可证

MIT
`, `
# Project Name

One sentence on what this project does and who it is for.

## Features

- Feature one: what it does and which problem it solves
- Feature two
- Feature three

## Quick start

\`\`\`bash
git clone https://example.com/org/project.git
cd project
npm install
npm run dev
\`\`\`

Open http://localhost:3000 to see it running.

## Configuration

| Variable | Purpose | Default |
| --- | --- | --- |
| \`PORT\` | Server port | \`3000\` |
| \`DATABASE_URL\` | Database connection string | none |
| \`LOG_LEVEL\` | Log level | \`info\` |

## Layout

\`\`\`
src/        source code
tests/      tests
docs/       documentation
scripts/    build and release scripts
\`\`\`

## Contributing

1. Fork the repository and create a branch
2. Run \`npm test\` before you commit
3. Open a pull request that explains the change and why

## License

MIT
`),
  },
];
