// PowerPoint templates for the new-file gallery (see lib.mjs): 16:9 decks drawn on blank slides, lengths in cm (33.867 × 19.05).
// Text sits in plain text boxes, left-aligned: Writer's slide editor and PowerPoint then place it the same way.
// A deck is a theme (colours) and four to six slides made of the pieces below: covers, a content page, cards, columns, steps…
import { esc, SW, SH, NAVY, CLAY, SLATE, PLUM, INK, OCHRE } from './lib.mjs';

export const cats = [['汇报总结', 'Reports'], ['商务营销', 'Business'], ['教育培训', 'Education'], ['简约主题', 'Minimal Themes']];

const b = s => `<b>${esc(s)}</b>`;
const W = 'FFFFFF', X0 = 2.4, CW = SW - 2 * X0; // the text margin and the content width
const light = P => Object.assign({ ink: '1D1D1F', bg: W, light: 'D5DEE9' }, P);

// ---- pieces ----
/** A content slide: its title (and a bar at its left), an optional subtitle, and the page number bottom right. */
async function page(p, C, n, title, sub) {
  const s = await p.slide(C.bg);
  await s.shape(X0, 1.55, 0.18, 1.3, C.acc);
  await s.text(X0 + 0.5, 1.3, 26, 1.6, b(title), { size: 26, color: C.ink });
  if (sub) await s.text(X0 + 0.5, 2.85, 26, 1, esc(sub), { size: 13, color: C.sub });
  await s.text(30.2, 17.5, 1.6, 0.8, String(n).padStart(2, '0'), { size: 10, color: C.sub });
  return s;
}
/** Cover with a full-height bar at the left and a soft panel on the right that carries a big mark (a quarter, a year). */
async function coverBand(p, C, { kicker, title, sub, mark }) {
  const s = await p.slide(C.bg);
  await s.shape(0, 0, 0.9, SH, C.acc);
  await s.shape(22.6, 0, SW - 22.6, SH, C.soft);
  if (mark) await s.text(23.4, 10.6, 10, 5, b(mark), { size: 120, color: C.light });
  await s.shape(2.9, 5.3, 1.6, 0.12, C.acc);
  await s.text(2.7, 5.8, 18, 1.2, esc(kicker), { size: 16, color: C.acc });
  await s.text(2.7, 7.1, 19, 3.6, b(title), { size: 40, color: C.ink });
  await s.text(2.7, 12.6, 18, 1.4, esc(sub), { size: 14, color: C.sub });
  return s;
}
/** Cover split in two: the left half in the accent colour carrying the title, the details on the right. */
async function coverSplit(p, C, { kicker, title, sub, lines }) {
  const s = await p.slide(C.bg);
  await s.shape(0, 0, 17.5, SH, C.acc);
  await s.text(2.4, 5.2, 13.6, 1.2, esc(kicker), { size: 14, color: C.light });
  await s.text(2.4, 6.6, 13.6, 6, b(title), { size: 40, color: W });
  await s.shape(2.4, 13.2, 1.6, 0.12, C.light);
  await s.text(2.4, 13.7, 13.6, 1.4, esc(sub), { size: 14, color: C.light });
  await s.text(19.6, 6.6, 12, 8, lines.map(l => ({ html: esc(l) })), { size: 14, color: C.sub });
  return s;
}
/** Cover with nothing but type: a thin rule, the title, the details. */
async function coverPlain(p, C, { kicker, title, sub }) {
  const s = await p.slide(C.bg);
  await s.shape(X0, 6.0, 1.6, 0.12, C.acc);
  await s.text(X0 - 0.2, 6.5, 26, 1.2, esc(kicker), { size: 14, color: C.sub });
  await s.text(X0 - 0.2, 7.8, 28, 4, b(title), { size: 40, color: C.ink });
  await s.text(X0 - 0.2, 13.2, 28, 1.6, esc(sub), { size: 14, color: C.sub });
  return s;
}
/** Cards with a big value, a label and a line of detail. */
async function cards(s, C, items, { y = 5.2, h = 7.6, fill } = {}) {
  const n = items.length, gap = 0.7, w = (CW - gap * (n - 1)) / n;
  for (const [i, [v, k, d]] of items.entries()) {
    const x = X0 + i * (w + gap);
    await s.shape(x, y, w, h, fill || C.soft, { geometry: 'roundRect' });
    await s.text(x + 0.8, y + 1.0, w - 1.6, 2.4, b(v), { size: n > 3 ? 30 : 40, color: C.acc });
    await s.text(x + 0.8, y + 3.9, w - 1.6, 1.1, b(k), { size: 16, color: C.ink });
    if (d) await s.text(x + 0.8, y + 5.1, w - 1.6, h - 5.3, esc(d), { size: 12, color: C.sub });
  }
}
/** Two columns, each a head over a list; number: true for a numbered list. */
async function twoCol(s, C, left, right, y = 5.3) {
  const col = async (x, { head, items, number }) => {
    await s.shape(x, y, 14, 0.06, C.acc);
    await s.text(x - 0.2, y + 0.4, 14, 1.2, b(head), { size: 18, color: C.acc });
    await s.text(x - 0.2, y + 1.9, 14, 9.5, items.map(h => ({ html: esc(h), list: number ? 'number' : 'bullet' })), { size: 15, color: C.ink });
  };
  await col(X0, left); await col(X0 + 15, right);
}
/** Numbered steps along a line, a title and a description under each. */
async function steps(s, C, items, y = 7.6) {
  const n = items.length, w = CW / n;
  await s.shape(X0 + w / 2, y + 0.55, CW - w, 0.06, C.line);
  for (const [i, [head, desc]] of items.entries()) {
    const x = X0 + i * w, cx = x + w / 2;
    await s.text(cx - 0.65, y - 0.05, 1.3, 1.3, b(String(i + 1)), { size: 13, color: W, fill: C.acc, geometry: 'ellipse', align: 'center' });
    await s.text(x + 0.3, y + 1.8, w - 0.6, 1.2, b(head), { size: 15, color: C.ink, align: 'center' });
    await s.text(x + 0.3, y + 3.1, w - 0.6, 3.6, esc(desc), { size: 12, color: C.sub, align: 'center' });
  }
}
/** Progress bars: name and note on the left, the bar and its percentage on the right. */
async function bars(s, C, items, y0 = 5.4) {
  for (const [i, [name, v, note]] of items.entries()) {
    const y = y0 + i * 3.6;
    await s.text(X0, y, 9, 1.2, b(name), { size: 18, color: C.ink });
    await s.text(X0, y + 1.3, 12, 1, esc(note), { size: 12, color: C.sub });
    await s.shape(14.6, y + 0.45, 14.4, 0.5, C.soft, { geometry: 'roundRect' });
    await s.shape(14.6, y + 0.45, 14.4 * v, 0.5, C.acc, { geometry: 'roundRect' });
    await s.text(29.3, y + 0.05, 2.6, 1.2, b(Math.round(v * 100) + '%'), { size: 16, color: C.acc });
  }
}
const table = (s, C, rows, y = 5.2) => s.table(rows, { x: X0, y, w: CW, h: Math.min(11.5, 1.15 * rows.length), header: C.acc, size: 14, color: C.ink, zebra: C.soft });
const bullets = (s, C, items, { x = X0, y = 5.3, w = CW, h = 10, size = 16, number } = {}) => s.text(x, y, w, h, items.map(h => ({ html: esc(h), list: number ? 'number' : 'bullet' })), { size, color: C.ink });
/** One big sentence with room around it, and who said it or what it means. */
async function statement(p, C, text, who) {
  const s = await p.slide(C.bg);
  await s.shape(4.4, 5.4, 1.6, 0.12, C.acc);
  await s.text(4.2, 6.0, 25.5, 6, b(text), { size: 32, color: C.ink });
  if (who) await s.text(4.2, 12.6, 25, 1.2, esc(who), { size: 14, color: C.sub });
  return s;
}
async function closing(p, C, title, sub) {
  const s = await p.slide(C.dark || C.acc);
  await s.shape(2.9, 7.3, 1.6, 0.12, C.light);
  await s.text(2.7, 7.8, 24, 2.6, b(title), { size: 48, color: W });
  if (sub) await s.text(2.7, 10.6, 24, 1.4, esc(sub), { size: 16, color: C.light });
  return s;
}

export default [
  {
    id: 'work-report', cat: '汇报总结', name: ['工作汇报', 'Work Report'],
    async build(p, t) {
      const C = light(NAVY);
      let s = await coverBand(p, C, { kicker: t('2026 年第三季度', 'Third quarter, 2026'), title: t('市场部工作汇报', 'Marketing Quarterly Review'), sub: t('汇报人：林晓　　2026 年 10 月 9 日', 'Alex Chen · October 9, 2026'), mark: 'Q3' });
      s = await page(p, C, 2, t('目录', 'Agenda'));
      const agenda = [[t('季度概览', 'Overview'), t('目标完成与整体表现', 'Targets and results')], [t('核心数据', 'Key numbers'), t('增长、转化与投入产出', 'Growth, conversion, return')], [t('重点项目', 'Projects'), t('三个项目的进展与问题', 'Where the three projects stand')], [t('下季度计划', 'Next quarter'), t('目标、举措与所需支持', 'Goals, actions, support needed')]];
      for (const [i, [h, d]] of agenda.entries()) {
        const x = X0 + i * 7.4;
        await s.shape(x, 7.2, 6.6, 0.06, C.line);
        await s.text(x - 0.2, 7.6, 6.6, 1.6, b('0' + (i + 1)), { size: 30, color: C.acc });
        await s.text(x - 0.2, 9.4, 6.6, 1.2, b(h), { size: 18, color: C.ink });
        await s.text(x - 0.2, 10.7, 6.6, 1.6, esc(d), { size: 12, color: C.sub });
      }
      s = await page(p, C, 3, t('核心数据', 'Key numbers'), t('第三季度主要指标完成情况', 'How the main targets landed this quarter'));
      await cards(s, C, [[t('536 万', '5.36M'), t('内容曝光', 'Content reach'), t('完成目标的 107%', '107% of target')], ['32%', t('会员复购率', 'Repeat purchase rate'), t('同比提升 4 个百分点', 'Up 4 points year on year')], ['1 : 4.8', t('投入产出比', 'Return on spend'), t('行业均值为 1 : 3.5', 'Industry average is 1 : 3.5')]]);
      await s.text(X0, 14.2, 29, 1, t('数据来源：各平台后台与会员系统，统计截至 9 月 30 日。', 'Source: platform dashboards and the member system, to September 30.'), { size: 11, color: C.sub });
      s = await page(p, C, 4, t('重点项目进展', 'Project progress'));
      await bars(s, C, [[t('秋季新品上市', 'Autumn product launch'), 0.9, t('物料全部到位，首发活动 10 月 1 日上线', 'All materials in place; launch event goes live October 1')], [t('会员体系升级', 'Membership upgrade'), 0.65, t('积分规则已定稿，系统开发进行中', 'Points rules final; system work under way')], [t('渠道拓展', 'New channels'), 0.4, t('两家渠道在谈，预计 11 月签约', 'Two partners in talks; signing expected in November')]]);
      s = await page(p, C, 5, t('下季度计划', 'Next quarter'));
      await twoCol(s, C, { head: t('目标', 'Goals'), items: [t('四季度内容曝光达到 600 万次', 'Reach 6 million in Q4'), t('会员复购率提升至 35%', 'Lift repeat purchases to 35%'), t('新增两家渠道合作伙伴', 'Sign two new channel partners')] },
        { head: t('需要的支持', 'Support needed'), items: [t('设计部支持双十一视觉物料', 'Design team for the November sale visuals'), t('IT 部排期会员系统二期开发', 'IT time for phase two of the member system'), t('追加渠道预算 20 万元', 'An extra $28,000 of channel budget')] });
      await closing(p, C, t('谢谢', 'Thank you'), t('欢迎提问与交流', 'Questions and discussion'));
    },
  },
  {
    id: 'year-review', cat: '汇报总结', name: ['年终总结', 'Year in Review'],
    async build(p, t) {
      const C = light({ acc: '1D1D1F', soft: 'F2F2F4', line: 'D2D2D7', sub: '6E6E73', light: 'B08D57' });
      let s = await p.slide('1D1D1F');
      await s.text(15.2, 2.6, 18, 8, b('2026'), { size: 150, color: '3A3A3C' });
      await s.shape(2.9, 9.6, 1.6, 0.12, C.light);
      await s.text(2.7, 10.1, 20, 1.2, t('产品部 · 年度工作总结', 'Product · Year in Review'), { size: 16, color: C.light });
      await s.text(2.7, 11.4, 24, 3.2, b(t('把一年做成一页', 'A year on one page')), { size: 44, color: W });
      await s.text(2.7, 15.2, 20, 1.2, t('李明　2026 年 12 月 20 日', 'Alex Chen · December 20, 2026'), { size: 14, color: 'A1A1AA' });
      s = await page(p, C, 2, t('年度关键数字', 'The year in numbers'));
      await cards(s, C, [['4', t('产品大版本', 'Major releases'), t('按季度节奏交付，零延期', 'One a quarter, none late')], [t('12 万', '120k'), t('月活跃用户', 'Monthly active users'), t('同比增长 65%', 'Up 65% year on year')], ['98.5%', t('服务可用性', 'Availability'), t('3 次故障，均在 30 分钟内恢复', 'Three incidents, all resolved within 30 minutes')]]);
      s = await page(p, C, 3, t('重点工作回顾', 'What we shipped'), t('四个季度，四件大事', 'Four quarters, four big things'));
      await steps(s, C, [[t('Q1 协作模块', 'Q1 · Collaboration'), t('实时协同上线，首月 2 万团队使用', 'Real-time editing launched; 20,000 teams in the first month')], [t('Q2 会员体系', 'Q2 · Membership'), t('积分与等级重构，续费率 +9 点', 'Points and tiers rebuilt; renewals up 9 points')], [t('Q3 智能纪要', 'Q3 · Smart notes'), t('会议转写与要点提炼，覆盖 80% 例会', 'Transcripts and summaries for 80% of recurring meetings')], [t('Q4 稳定性专项', 'Q4 · Reliability'), t('打开速度快 40%，故障时长减半', 'Opens 40% faster; incident time halved')]]);
      s = await page(p, C, 4, t('收获与不足', 'What worked, what did not'));
      await twoCol(s, C, { head: t('做得好的', 'Worked'), items: [t('版本节奏稳定，跨部门配合顺畅', 'A steady release rhythm and smooth cross-team work'), t('用户反馈闭环建立，响应周期缩短到 2 周', 'The feedback loop is in place; turnaround down to two weeks'), t('新人培养见效，3 人独立带项目', 'Three new hires now run projects on their own')] },
        { head: t('需要改进的', 'Needs work'), items: [t('需求评审偏晚，二季度返工较多', 'Reviews came late; Q2 had too much rework'), t('数据埋点不全，部分决策靠经验', 'Gaps in analytics left some decisions to gut feel'), t('文档沉淀不足，知识集中在少数人', 'Too little written down; knowledge sits with a few people')] });
      s = await page(p, C, 5, t('2027 年计划', 'Plan for 2027'));
      await table(s, C, [[t('目标', 'Goal'), t('关键举措', 'Key actions'), t('时间', 'When'), t('负责人', 'Owner')],
        [t('月活达到 20 万', '200k monthly actives'), t('移动端重构、邀请机制', 'Mobile rebuild, invitations'), t('全年', 'All year'), t('李明', 'Alex Chen')],
        [t('付费转化率 6%', '6% paid conversion'), t('团队版功能分层与试用', 'Team tiers and trials'), 'Q1–Q2', t('王芳', 'Jordan Lee')],
        [t('可用性 99.9%', '99.9% availability'), t('多活部署、变更审批', 'Multi-region deployment, change control'), 'Q1–Q3', t('刘洋', 'Casey Kim')],
        [t('知识库覆盖全部模块', 'Docs for every module'), t('每个版本附带文档更新', 'Docs shipped with every release'), t('持续', 'Ongoing'), t('陈静', 'Morgan Reed')]]);
      await closing(p, Object.assign({}, C, { dark: '1D1D1F' }), t('感谢一路同行', 'Thank you for the year'), t('2027，继续前进', 'On to 2027'));
    },
  },
  {
    id: 'business-plan', cat: '商务营销', name: ['商业计划书', 'Business Plan'],
    async build(p, t) {
      const C = light(Object.assign({}, PLUM, { light: 'D7D0E1' }));
      let s = await coverSplit(p, C, { kicker: t('商业计划书 · 2026', 'Business plan · 2026'), title: t('「一起写」协作文档平台', '“Together” collaboration platform'), sub: t('面向中小团队的轻量协作工具', 'Lightweight collaboration for small teams'), lines: [t('远山科技有限公司', 'Brightline Technologies'), t('创始人：李明', 'Founder: Alex Chen'), 'hello@example.com', '', t('本文件仅供投资人内部参考', 'For prospective investors only')] });
      s = await page(p, C, 2, t('问题与机会', 'Problem and opportunity'));
      await twoCol(s, C, { head: t('中小团队的痛点', 'What small teams struggle with'), items: [t('协作工具太重：功能多，用起来的不到两成', 'Tools built for enterprises: teams use less than a fifth of the features'), t('文档、表格、演示分散在三四个产品里', 'Documents, sheets and slides live in three or four products'), t('版本靠聊天软件传来传去，返工多', 'Versions travel through chat, and rework follows')] },
        { head: t('为什么是现在', 'Why now'), items: [t('中小团队数字化渗透率不足 30%，仍在快速增长', 'Under 30% of small teams have adopted such tools, and adoption is growing fast'), t('订阅付费习惯已经形成', 'Paying by subscription is now normal'), t('AI 大幅降低了内容起草与整理的门槛', 'AI has cut the cost of drafting and organising content')] });
      s = await page(p, C, 3, t('产品与解决方案', 'The product'), t('一处编辑，随手协作，AI 起草', 'One place to write, collaborate and draft with AI'));
      await cards(s, C, [[t('一处', 'One'), t('一个地方写完所有东西', 'One place for everything'), t('文档、表格、演示、导图同一入口，格式与桌面软件互通', 'Documents, sheets, slides and mind maps behind one door, in formats desktop software opens')], [t('3 秒', '3 s'), t('分享即协作', 'Sharing is collaborating'), t('一个链接进入，多人实时编辑，改动可追溯', 'One link in, real-time editing for everyone, every change traceable')], [t('10×', '10×'), t('AI 起草与整理', 'Drafting with AI'), t('从一句话到初稿，从录音到纪要', 'From a sentence to a first draft, from a recording to minutes')]]);
      s = await page(p, C, 4, t('市场规模', 'Market'));
      await bars(s, C, [[t('可触达市场', 'Addressable market'), 1, t('全国中小团队协作软件支出约 180 亿元 / 年', 'About 18 billion a year spent by small teams on collaboration software')], [t('可服务市场', 'Serviceable market'), 0.4, t('知识型团队，72 亿元 / 年', 'Knowledge-work teams: 7.2 billion a year')], [t('三年目标', 'Three-year target'), 0.05, t('9 亿元营收，约 1.2% 份额', '900 million in revenue, about a 1.2% share')]]);
      s = await page(p, C, 5, t('商业模式与财务预测', 'Business model and forecast'), t('免费个人版获客，团队版按席位订阅', 'A free personal tier brings people in; teams pay per seat'));
      await table(s, C, [[t('年份', 'Year'), t('付费团队', 'Paying teams'), t('年经常性收入', 'Annual recurring revenue'), t('毛利率', 'Gross margin'), t('团队规模', 'Headcount')],
        ['2027', '2,000', t('1,400 万元', '14M'), '68%', '28'], ['2028', '9,000', t('7,200 万元', '72M'), '74%', '65'], ['2029', '30,000', t('2.6 亿元', '260M'), '78%', '140']]);
      await s.text(X0, 10.6, CW, 2, t('定价：个人版免费；团队版每席位每月 29 元，年付八折；企业版按需报价。获客成本回收期目标 9 个月。', 'Pricing: personal free; team 29 per seat per month, 20% off annually; enterprise on request. Target payback on acquisition cost: nine months.'), { size: 13, color: C.sub });
      s = await page(p, C, 6, t('团队与融资', 'Team and the round'));
      await twoCol(s, C, { head: t('核心团队', 'Team'), items: [t('李明，CEO：10 年协作产品经验，前大型办公软件产品负责人', 'Alex Chen, CEO: ten years in collaboration products, formerly led product at a major office suite'), t('王芳，CTO：分布式系统与实时协同技术', 'Jordan Lee, CTO: distributed systems and real-time sync'), t('张伟，COO：中小企业渠道与销售', 'Sam Taylor, COO: small-business channels and sales')] },
        { head: t('本轮融资', 'This round'), items: [t('融资 3,000 万元，出让 15% 股份', 'Raising 30 million for 15%'), t('用途：研发 55%、市场 30%、运营 15%', 'Use of funds: 55% engineering, 30% marketing, 15% operations'), t('里程碑：18 个月内达到 9,000 付费团队', 'Milestone: 9,000 paying teams within 18 months')] });
    },
  },
  {
    id: 'product-launch', cat: '商务营销', name: ['产品发布', 'Product Launch'],
    async build(p, t) {
      const C = { acc: 'E0B25A', soft: '24262C', line: '35383F', sub: 'A1A1AA', ink: 'F5F5F7', bg: '15171B', light: 'F0DCA8', dark: '15171B' };
      let s = await p.slide(C.bg);
      await s.text(X0 - 0.2, 5.6, 20, 1.2, t('2026 秋季发布', 'Autumn 2026 launch'), { size: 16, color: C.acc });
      await s.text(X0 - 0.2, 7.0, 28, 4.6, b(t('一起写 3.0', 'Together 3.0')), { size: 72, color: C.ink });
      await s.text(X0 - 0.2, 12.2, 26, 1.6, t('让工具退到工作背后', 'The tool steps back. The work steps forward.'), { size: 20, color: C.sub });
      await s.shape(X0, 15.2, 1.6, 0.12, C.acc);
      s = await statement(p, C, t('每一次协作，都不该从「发我一下最新版」开始。', 'No collaboration should start with “send me the latest version”.'), t('我们问了 1,200 个团队，这是他们最常说的一句话。', 'We asked 1,200 teams. This is the sentence they said most.'));
      s = await page(p, C, 3, t('三个变化', 'Three things that changed'));
      await cards(s, C, [[t('实时', 'Live'), t('多人同时编辑', 'Everyone edits at once'), t('光标、选区、改动即时可见，冲突自动合并', 'Cursors, selections and edits show instantly; conflicts merge on their own')], [t('5 分钟', '5 min'), t('会议变纪要', 'Meetings become minutes'), t('录音转写、要点提炼、待办分发，一次完成', 'Transcribe, summarise and assign the actions in one pass')], [t('离线', 'Offline'), t('没有网络也能写', 'Works without a connection'), t('本地保存，联网后自动同步', 'Saved locally, synced when you are back')]]);
      s = await page(p, C, 4, t('与上一代相比', 'Compared with 2.0'));
      await s.table([[t('项目', ''), '2.0', '3.0'], [t('打开 100 页文档', 'Opening a 100-page document'), t('3.2 秒', '3.2 s'), t('0.8 秒', '0.8 s')], [t('同时编辑人数', 'Simultaneous editors'), '10', '100'], [t('离线编辑', 'Offline editing'), t('不支持', 'No'), t('支持', 'Yes')], [t('AI 起草与纪要', 'AI drafting and minutes'), t('部分', 'Partial'), t('全部格式', 'Every format')]], { x: X0, y: 5.4, w: CW, h: 6.5, header: C.soft, size: 15, color: C.ink });
      s = await page(p, C, 5, t('价格与上市时间', 'Price and availability'));
      await cards(s, C, [[t('¥0', '$0'), t('个人版', 'Personal'), t('全部功能，不限文档数', 'Every feature, unlimited documents')], [t('¥29', '$4'), t('团队版 · 每人每月', 'Team · per person per month'), t('共享空间、权限与审计', 'Shared spaces, permissions and audit')], [t('10 月 15 日', 'October 15'), t('全平台上线', 'On every platform'), t('Mac、Windows、网页同步发布', 'Mac, Windows and the web on the same day')]]);
      await closing(p, C, t('今天就开始', 'Start today'), 'example.com/together');
    },
  },
  {
    id: 'lesson', cat: '教育培训', name: ['课件', 'Lesson Slides'],
    async build(p, t) {
      const C = light(SLATE);
      let s = await coverBand(p, C, { kicker: t('高中物理 · 必修一', 'Physics · Year 11'), title: t('牛顿第二定律', 'Newton’s Second Law'), sub: t('第三章 第 2 课时　授课：王老师', 'Chapter 3, lesson 2 · Ms. Lee'), mark: 'F=ma' });
      s = await page(p, C, 2, t('学习目标', 'What we will learn'));
      await bullets(s, C, [t('理解加速度与力、质量之间的关系，能用自己的话说出牛顿第二定律。', 'Explain in your own words how acceleration depends on force and mass.'), t('掌握 F = ma 的含义与单位，会判断公式中每个量的方向。', 'Know what each quantity in F = ma means, its unit and its direction.'), t('能对简单情境做受力分析并列式求解。', 'Draw a force diagram for a simple situation and solve it.'), t('体会「控制变量」的实验方法。', 'See how a controlled experiment isolates one variable at a time.')], { number: true, size: 18 });
      s = await page(p, C, 3, t('新知讲解', 'The idea'), t('从实验到公式', 'From the experiment to the formula'));
      await twoCol(s, C, { head: t('定律内容', 'The law'), items: [t('物体的加速度与所受合外力成正比，与质量成反比。', 'The acceleration of an object is proportional to the net force on it and inversely proportional to its mass.'), t('公式：F = ma，单位：牛顿（N）= kg·m/s²', 'F = ma, with the newton (N) = kg·m/s²'), t('加速度的方向与合外力方向相同。', 'The acceleration points the same way as the net force.')] },
        { head: t('理解要点', 'Things to notice'), items: [t('F 指合外力，不是某一个力。', 'F is the net force, not any single force.'), t('质量越大，同样的力产生的加速度越小。', 'The more mass, the less acceleration the same force gives.'), t('力与加速度同时存在、同时变化。', 'Force and acceleration appear, change and vanish together.')] });
      s = await page(p, C, 4, t('例题', 'Worked example'), t('质量 2 kg 的物体在光滑水平面上受 6 N 的水平拉力，求加速度。', 'A 2 kg block on a smooth table is pulled horizontally with 6 N. Find its acceleration.'));
      await steps(s, C, [[t('审题', 'Read'), t('找出已知量：m = 2 kg，F = 6 N；光滑说明无摩擦', 'Known: m = 2 kg, F = 6 N; “smooth” means no friction')], [t('受力分析', 'Forces'), t('重力与支持力抵消，合外力即拉力', 'Weight and the normal force cancel; the net force is the pull')], [t('列方程', 'Equation'), t('由 F = ma 得 a = F / m', 'From F = ma, a = F / m')], [t('求解', 'Solve'), t('a = 6 / 2 = 3 m/s²，方向与拉力相同', 'a = 6 / 2 = 3 m/s², in the direction of the pull')]]);
      s = await page(p, C, 5, t('课堂练习', 'Practice'));
      await bullets(s, C, [t('一辆 1,000 kg 的小车在 2,000 N 的牵引力下启动，阻力 500 N，求加速度。', 'A 1,000 kg car starts with a 2,000 N driving force against 500 N of resistance. Find its acceleration.'), t('同样的力作用在质量为 m 和 2m 的物体上，加速度之比是多少？', 'The same force acts on masses m and 2m. What is the ratio of their accelerations?'), t('一个物体做匀速直线运动，它受到的合外力是多少？说明理由。', 'An object moves in a straight line at constant speed. What is the net force on it, and why?')], { number: true, size: 17 });
      s = await page(p, C, 6, t('小结与作业', 'Summary and homework'));
      await twoCol(s, C, { head: t('本课小结', 'Summary'), items: [t('F = ma：合外力决定加速度', 'F = ma: the net force sets the acceleration'), t('先受力分析，再列式', 'Forces first, equation second'), t('注意单位与方向', 'Watch the units and the directions')] },
        { head: t('课后作业', 'Homework'), items: [t('课本第 78 页第 3、5、7 题', 'Textbook p. 78, questions 3, 5 and 7'), t('预习：牛顿第三定律', 'Read ahead: Newton’s third law'), t('选做：设计一个验证 F = ma 的小实验', 'Optional: design a small experiment to test F = ma')] });
    },
  },
  {
    id: 'thesis-defense', cat: '教育培训', name: ['论文答辩', 'Thesis Defence'],
    async build(p, t) {
      const C = light(NAVY);
      let s = await coverPlain(p, C, { kicker: t('本科毕业论文答辩 · 江城大学管理学院', 'Undergraduate thesis defence · Lakeside University'), title: t('基于用户行为数据的协作工具留存分析', 'Retention in Collaboration Tools: Evidence from Behavioural Data'), sub: t('答辩人：李明　　指导教师：王芳 教授　　2026 年 5 月 20 日', 'Alex Chen · Supervisor: Prof. Jordan Lee · May 20, 2026') });
      s = await page(p, C, 2, t('研究背景与问题', 'Background and questions'));
      await twoCol(s, C, { head: t('研究背景', 'Background'), items: [t('协作工具用户获取成本高，但次月流失普遍超过 50%', 'Collaboration tools pay dearly for users, yet most lose over half of them by the second month'), t('已有研究多用漏斗描述流失，少有早期行为的量化分析', 'Prior work describes churn with funnels; few studies quantify early behaviour'), t('产品团队需要可操作的引导策略依据', 'Product teams need evidence they can act on')] },
        { head: t('研究问题', 'Questions'), number: true, items: [t('哪些首周行为与次月留存显著相关？', 'Which first-week behaviours predict second-month retention?'), t('协作规模对留存的作用是否存在阈值？', 'Does the effect of collaboration size have a threshold?'), t('发现如何转化为新用户引导策略？', 'How do the findings translate into onboarding?')] });
      s = await page(p, C, 3, t('研究方法', 'Method'), t('48,000 名用户 · 2025 年 1–12 月脱敏行为日志', '48,000 users · anonymised logs, January–December 2025'));
      await steps(s, C, [[t('数据获取', 'Data'), t('注册后 8 周的事件日志，清洗后 46,812 人', 'Eight weeks of events per user; 46,812 after cleaning')], [t('变量构建', 'Variables'), t('首周协作人数、功能广度与深度、设备类型', 'First-week collaborators, feature breadth and depth, device')], [t('模型估计', 'Model'), t('逻辑回归 + 分组对比，控制注册渠道', 'Logistic regression and cohort comparison, controlling for channel')], [t('稳健性检验', 'Robustness'), t('更换留存窗口与样本区间复核', 'Re-run with other retention windows and sample periods')]]);
      s = await page(p, C, 4, t('主要结果', 'Results'));
      await table(s, C, [[t('变量', 'Variable'), t('系数', 'Coefficient'), t('显著性', 'p'), t('解读', 'Reading')],
        [t('首周完成一次多人协作', 'One multi-person collaboration in week 1'), '+0.82', '< 0.01', t('次月留存率提高 19 个百分点', 'Second-month retention up 19 points')],
        [t('功能使用广度（每增 1 项）', 'Feature breadth (per feature)'), '+0.11', '< 0.05', t('边际效应递减，5 项后不显著', 'Diminishing; not significant beyond five')],
        [t('协作人数 ≥ 3', 'Three or more collaborators'), '+0.35', '< 0.01', t('存在阈值：3 人以上增益趋平', 'A threshold: gains flatten past three')],
        [t('仅移动端使用', 'Mobile only'), '−0.27', '< 0.05', t('留存显著低于桌面端用户', 'Retains markedly worse than desktop users')]]);
      s = await page(p, C, 5, t('结论与展望', 'Conclusions and further work'));
      await twoCol(s, C, { head: t('主要结论', 'Conclusions'), items: [t('首周的一次真实协作是最强的留存信号', 'One real collaboration in the first week is the strongest retention signal'), t('引导应聚焦「邀请一位同事」，而非功能巡礼', 'Onboarding should push “invite a colleague”, not a feature tour'), t('移动端首次体验需要单独设计', 'The first mobile session deserves its own design')] },
        { head: t('不足与展望', 'Limits and next steps'), items: [t('单一产品数据，外推需谨慎', 'One product’s data; generalise with care'), t('相关性分析，因果有待实验验证', 'Correlational; causation awaits an experiment'), t('后续可加入文本与会话内容特征', 'Text and session features could be added next')] });
      await closing(p, C, t('请各位老师批评指正', 'Thank you'), t('谢谢！', 'Questions and comments are welcome'));
    },
  },
  {
    id: 'training', cat: '教育培训', name: ['培训', 'Training'],
    async build(p, t) {
      const C = light(Object.assign({}, CLAY, { light: 'F0D9CC' }));
      let s = await coverBand(p, C, { kicker: t('新员工入职培训 · 第一天', 'New-hire orientation · Day one'), title: t('欢迎加入远山科技', 'Welcome to Brightline'), sub: t('人力资源部　2026 年 10 月 12 日', 'Human Resources · October 12, 2026'), mark: t('你好', 'Hi') });
      s = await page(p, C, 2, t('今日安排', 'Today'));
      await table(s, C, [[t('时间', 'Time'), t('内容', 'Session'), t('讲师', 'Led by')], ['09:00 – 09:45', t('公司介绍与文化', 'The company and how we work'), t('张伟 · 总经理', 'Sam Taylor, General Manager')], ['10:00 – 11:30', t('制度与流程：考勤、报销、审批', 'Policies: attendance, expenses, approvals'), t('陈静 · 人力资源', 'Morgan Reed, HR')], ['13:30 – 15:00', t('工具与账号：邮箱、协作文档、内部系统', 'Tools: mail, shared documents, internal systems'), t('刘洋 · IT', 'Casey Kim, IT')], ['15:15 – 16:30', t('信息安全与合规', 'Security and compliance'), t('王芳 · 法务', 'Jordan Lee, Legal')], ['16:30 – 17:00', t('答疑与反馈', 'Questions and feedback'), t('人力资源部', 'HR')]]);
      s = await page(p, C, 3, t('培训目标', 'By the end of today'));
      await cards(s, C, [[t('了解', 'Know'), t('公司与团队', 'The company and your team'), t('我们做什么、为谁做、怎么协作', 'What we make, for whom, and how we work together')], [t('熟悉', 'Find'), t('制度与流程', 'The policies'), t('知道规则在哪里、找谁问', 'Where the rules live and whom to ask')], [t('掌握', 'Use'), t('日常工具', 'The everyday tools'), t('今天下班前，所有账号可用', 'Every account working before you leave today')]]);
      s = await page(p, C, 4, t('核心内容', 'The essentials'));
      await twoCol(s, C, { head: t('公司与文化', 'Company and culture'), items: [t('三条价值观：用户第一、坦诚沟通、把事做完', 'Three values: users first, candour, finish what you start'), t('组织结构与各部门的职责', 'Who does what across the organisation'), t('沟通习惯：文档先行，会议有结论', 'Habits: write it down first; every meeting ends with a decision')] },
        { head: t('制度与流程', 'Policies and process'), items: [t('工作时间与考勤、请假申请', 'Hours, attendance and leave requests'), t('报销标准与审批路径', 'Expense rules and the approval path'), t('信息安全：设备、密码、数据分级', 'Security: devices, passwords, data classes')] });
      s = await page(p, C, 5, t('案例练习', 'Case exercise'), t('情境：客户在周五下午 5 点发来紧急需求，要求周一上线。', 'Scenario: a customer sends an urgent request at 5 pm Friday and wants it live on Monday.'));
      await steps(s, C, [[t('阅读场景', 'Read'), t('5 分钟，各自记下你会先做的三件事', 'Five minutes; note the three things you would do first')], [t('小组讨论', 'Discuss'), t('4 人一组，形成一个共同方案', 'Groups of four agree on one plan')], [t('分享', 'Share'), t('每组 2 分钟，说清做什么、不做什么', 'Two minutes per group: what you do and what you don’t')], [t('点评', 'Debrief'), t('讲师结合流程与价值观点评', 'The trainer ties it back to process and values')]]);
      await closing(p, C, t('有问题，随时找我们', 'Questions? Come find us'), t('培训反馈问卷将在今晚发到你的邮箱', 'A feedback form arrives in your inbox tonight'));
    },
  },
  {
    id: 'proposal', cat: '商务营销', name: ['提案', 'Proposal'],
    async build(p, t) {
      const C = light(Object.assign({}, OCHRE, { light: 'E3D7C0' }));
      let s = await coverPlain(p, C, { kicker: t('内部提案 · 产品部', 'Internal proposal · Product'), title: t('建立客户反馈闭环机制', 'A closed loop for customer feedback'), sub: t('提案人：李明　　2026 年 9 月 25 日', 'Alex Chen · September 25, 2026') });
      s = await page(p, C, 2, t('现状与问题', 'Where we are'));
      await s.text(X0, 5.4, 10, 3, b(t('6 周', '6 weeks')), { size: 64, color: C.acc });
      await s.text(X0, 9.0, 10, 2, t('一条高频需求从提出到有人响应的平均时间', 'The average wait from a frequent request to any response'), { size: 13, color: C.sub });
      await bullets(s, C, [t('反馈散落在客服工单、销售群和邮件里，没有统一入口', 'Feedback is scattered across support tickets, sales chats and email, with no single entry point'), t('没有固定的评审节奏，谁声音大谁优先', 'No review cadence; whoever shouts loudest goes first'), t('客户不知道问题有没有人在管，重复提、反复问', 'Customers cannot tell whether anyone is on it, so they ask again and again')], { x: 13.4, y: 5.4, w: 18, h: 9, size: 15 });
      s = await page(p, C, 3, t('提案方案', 'The proposal'), t('三件事，一个闭环', 'Three parts, one loop'));
      await cards(s, C, [[t('01', '01'), t('统一入口', 'One inbox'), t('所有渠道的反馈进入同一个看板，自动去重与归类', 'Feedback from every channel lands on one board, deduplicated and tagged')], [t('02', '02'), t('每周评审', 'Weekly review'), t('产品、客服、销售固定 45 分钟，给每条高频需求一个明确状态', 'Product, support and sales meet for 45 minutes and give every frequent request a status')], [t('03', '03'), t('公开路线图', 'Public roadmap'), t('客户能看到「已计划 / 进行中 / 已上线」，上线自动通知提出者', 'Customers see planned / in progress / shipped, and get told when it ships')]]);
      s = await page(p, C, 4, t('实施步骤', 'How we get there'));
      await steps(s, C, [[t('10 月上旬', 'Early October'), t('接入三个渠道，搭建反馈看板', 'Connect the three channels and set up the board')], [t('10 月中旬', 'Mid-October'), t('首次周评审，确定状态规则', 'First weekly review; agree the status rules')], [t('11 月', 'November'), t('公开路线图上线，自动通知打通', 'Public roadmap live; notifications wired up')], [t('12 月', 'December'), t('复盘：响应周期与客户满意度', 'Review: response time and customer satisfaction')]]);
      s = await page(p, C, 5, t('投入与收益', 'Cost and return'));
      await twoCol(s, C, { head: t('投入', 'Cost'), items: [t('工具：反馈看板与通知，约 2 万元 / 年', 'Tools: board and notifications, about 20,000 a year'), t('人力：产品 0.5 人，客服 0.3 人', 'People: half a product manager, a third of a support agent'), t('每周 45 分钟评审会', 'A 45-minute meeting every week')] },
        { head: t('预期收益', 'Return'), items: [t('响应周期从 6 周缩短到 2 周', 'Response time from six weeks to two'), t('重复反馈减少 40%，客服工单下降', 'Duplicate feedback down 40%; fewer tickets'), t('续费沟通中有据可依，目标续费率 +5 个百分点', 'Renewal conversations backed by evidence; target renewals up 5 points')] });
      s = await page(p, C, 6, t('需要的决策', 'Decisions needed'));
      await bullets(s, C, [t('同意成立由产品、客服、销售各一人组成的评审小组，每周固定时间。', 'Approve a review group of one person each from product, support and sales, meeting weekly.'), t('批准 2 万元 / 年的工具预算。', 'Approve the 20,000-a-year tool budget.'), t('授权产品部对外发布路线图的「已计划」状态。', 'Authorise Product to publish the “planned” status externally.')], { number: true, size: 18 });
      await s.text(X0, 13.6, CW, 1.5, t('如获批准，10 月 8 日启动，12 月底汇报第一阶段结果。', 'If approved, we start on October 8 and report on the first phase at the end of December.'), { size: 14, color: C.sub });
    },
  },
  {
    id: 'minimal-light', cat: '简约主题', name: ['极简 · 浅色', 'Minimal · Light'],
    async build(p, t) {
      const C = light(Object.assign({}, INK, { light: 'AEAEB2' }));
      let s = await p.slide(W);
      await s.text(X0 - 0.2, 6.4, 28, 4.2, b(t('把话说清楚', 'Say it plainly')), { size: 54, color: C.ink });
      await s.shape(X0, 11.2, 2.4, 0.08, C.ink);
      await s.text(X0 - 0.2, 11.8, 26, 1.4, t('一个不打扰内容的演示模板', 'A presentation template that stays out of the way'), { size: 16, color: C.sub });
      await s.text(X0 - 0.2, 16.6, 26, 1, t('远山科技 · 2026 年 10 月', 'Brightline · October 2026'), { size: 11, color: C.sub });
      s = await p.slide(W);
      await s.text(X0 - 0.2, 5.0, 10, 4, b('01'), { size: 80, color: 'E5E5EA' });
      await s.text(X0 - 0.2, 10.2, 26, 2, b(t('从问题开始', 'Start with the problem')), { size: 36, color: C.ink });
      await s.text(X0 - 0.2, 12.6, 26, 1.4, t('每一节只讲一件事', 'One idea per section'), { size: 16, color: C.sub });
      s = await p.slide(W);
      await s.text(X0 - 0.2, 2.6, 28, 1.8, b(t('三个要点', 'Three points')), { size: 30, color: C.ink });
      for (const [i, [h, d]] of [[t('少即是多', 'Less is more'), t('每页不超过三条信息，大字号，留白多。', 'No more than three things per slide, large type, plenty of air.')], [t('对齐一切', 'Align everything'), t('左边距固定，所有元素从同一条线开始。', 'One left margin; everything starts on the same line.')], [t('只用一种强调', 'One kind of emphasis'), t('加粗或颜色，选一种，全篇一致。', 'Bold or colour, pick one and keep it.')]].entries()) {
        const y = 5.6 + i * 3.6;
        await s.shape(X0, y + 0.55, 0.5, 0.08, C.ink);
        await s.text(X0 + 1.2, y, 27, 1.2, b(h), { size: 20, color: C.ink });
        await s.text(X0 + 1.2, y + 1.3, 27, 1.6, esc(d), { size: 14, color: C.sub });
      }
      s = await p.slide(W);
      await s.text(X0 - 0.2, 2.6, 28, 1.8, b(t('之前 · 之后', 'Before · After')), { size: 30, color: C.ink });
      await s.shape(16.9, 5.4, 0.06, 10, 'E5E5EA');
      await s.text(X0 - 0.2, 5.4, 13, 1.2, esc(t('之前', 'Before')), { size: 13, color: C.sub });
      await s.text(X0 - 0.2, 6.8, 13, 8, [t('每页七八条要点', 'Seven or eight bullets a slide'), t('三种字体，五种颜色', 'Three fonts, five colours'), t('图表挤在角落', 'Charts squeezed into corners')].map(h => ({ html: esc(h), list: 'bullet' })), { size: 16, color: C.ink });
      await s.text(18.1, 5.4, 13, 1.2, esc(t('之后', 'After')), { size: 13, color: C.sub });
      await s.text(18.1, 6.8, 13, 8, [t('一页一个结论', 'One conclusion per slide'), t('一种字体，黑白灰', 'One font, black, white and grey'), t('图表独占一页', 'A chart gets its own slide')].map(h => ({ html: esc(h), list: 'bullet' })), { size: 16, color: C.ink });
      await statement(p, C, t('「简单不是少，而是没有多余。」', '“Simple is not less. It is nothing extra.”'), t('—— 设计原则', '— A design principle'));
      s = await p.slide(W);
      await s.text(X0 - 0.2, 7.4, 28, 3, b(t('谢谢', 'Thank you')), { size: 54, color: C.ink });
      await s.shape(X0, 11.2, 2.4, 0.08, C.ink);
      await s.text(X0 - 0.2, 11.8, 26, 1.4, 'hello@example.com', { size: 16, color: C.sub });
    },
  },
  {
    id: 'minimal-dark', cat: '简约主题', name: ['极简 · 深色', 'Minimal · Dark'],
    async build(p, t) {
      const C = { acc: 'F5F5F7', soft: '1C1C1E', line: '2C2C2E', sub: '8E8E93', ink: 'F5F5F7', bg: '0F0F10', light: '8E8E93', dark: '0F0F10' };
      let s = await p.slide(C.bg);
      await s.text(X0 - 0.2, 6.4, 28, 4.2, b(t('把话说清楚', 'Say it plainly')), { size: 54, color: C.ink });
      await s.shape(X0, 11.2, 2.4, 0.08, C.ink);
      await s.text(X0 - 0.2, 11.8, 26, 1.4, t('一个不打扰内容的演示模板 · 深色', 'A presentation template that stays out of the way · dark'), { size: 16, color: C.sub });
      await s.text(X0 - 0.2, 16.6, 26, 1, t('远山科技 · 2026 年 10 月', 'Brightline · October 2026'), { size: 11, color: C.sub });
      s = await p.slide(C.bg);
      await s.text(X0 - 0.2, 5.0, 10, 4, b('01'), { size: 80, color: '2C2C2E' });
      await s.text(X0 - 0.2, 10.2, 26, 2, b(t('从问题开始', 'Start with the problem')), { size: 36, color: C.ink });
      await s.text(X0 - 0.2, 12.6, 26, 1.4, t('每一节只讲一件事', 'One idea per section'), { size: 16, color: C.sub });
      s = await p.slide(C.bg);
      await s.text(X0 - 0.2, 2.6, 28, 1.8, b(t('三个要点', 'Three points')), { size: 30, color: C.ink });
      await cards(s, C, [['01', t('少即是多', 'Less is more'), t('每页不超过三条信息，大字号，留白多。', 'No more than three things per slide, large type, plenty of air.')], ['02', t('对齐一切', 'Align everything'), t('左边距固定，所有元素从同一条线开始。', 'One left margin; everything starts on the same line.')], ['03', t('只用一种强调', 'One kind of emphasis'), t('加粗或颜色，选一种，全篇一致。', 'Bold or colour, pick one and keep it.')]], { y: 5.6, h: 8.2 });
      s = await p.slide(C.bg);
      await s.text(X0 - 0.2, 2.6, 28, 1.8, b(t('数字说话', 'Let the numbers talk')), { size: 30, color: C.ink });
      for (const [i, [v, k]] of [['3×', t('更快的加载', 'Faster loading')], ['−40%', t('更少的会议', 'Fewer meetings')], ['98%', t('满意度', 'Satisfaction')]].entries()) {
        const x = X0 + i * 9.9;
        await s.text(x - 0.2, 6.2, 9, 3, b(v), { size: 60, color: C.ink });
        await s.shape(x, 9.9, 0.8, 0.08, C.sub);
        await s.text(x - 0.2, 10.4, 9, 1.4, esc(k), { size: 15, color: C.sub });
      }
      await statement(p, C, t('「简单不是少，而是没有多余。」', '“Simple is not less. It is nothing extra.”'), t('—— 设计原则', '— A design principle'));
      s = await p.slide(C.bg);
      await s.text(X0 - 0.2, 7.4, 28, 3, b(t('谢谢', 'Thank you')), { size: 54, color: C.ink });
      await s.shape(X0, 11.2, 2.4, 0.08, C.ink);
      await s.text(X0 - 0.2, 11.8, 26, 1.4, 'hello@example.com', { size: 16, color: C.sub });
    },
  },
];
