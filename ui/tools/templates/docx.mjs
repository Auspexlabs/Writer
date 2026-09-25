// Word templates for the new-file gallery (see lib.mjs for how a template is written, build.mjs for how it is built).
import { span, esc, bold, lines, NAVY, INK, CLAY, SLATE, PLUM, OCHRE } from './lib.mjs';

export const cats = [['简历求职', 'Résumés & Jobs'], ['报告计划', 'Reports & Plans'], ['行政办公', 'Office & Admin'], ['合同信函', 'Contracts & Letters'], ['学习', 'Study']];

// ---- shared pieces ----
const heading = (d, C, text, level = 2) => d[level === 1 ? 'h1' : level === 3 ? 'h3' : 'h2'](span(text, { color: C.acc }));
/** A résumé or letter section: the heading, then a thin rule. */
const section = async (d, C, text) => { await d.h2(span(text, { color: C.acc })); await d.rule(C.line); };
/** Key/value pairs laid out two to a row, the keys on a soft fill. */
const infoTable = (d, C, pairs, widths = ['2.6cm', '5.4cm', '2.6cm', '5.4cm']) => {
  const rows = [];
  for (let i = 0; i < pairs.length; i += 2) rows.push(pairs.slice(i, i + 2).flatMap(([k, v]) => [{ html: span(k, { bold: true, color: C.acc }), fill: C.soft }, esc(v)]));
  return d.table(rows, { widths, borders: 'all', borderColor: C.line });
};
/** A header row on the accent colour, then the rows. */
const dataTable = (d, C, head, rows, o = {}) => d.table([head.map(h => ({ html: span(h, { bold: true, color: 'FFFFFF' }), fill: C.acc })), ...rows], Object.assign({ width: '100%', borders: 'horizontal', borderColor: C.line }, o));
/** One entry of a résumé: organisation and role on the left, the dates on the right, then its bullet points. */
const entry = async (d, C, org, role, when, items) => {
  await d.table([[lines(bold(org), span(role, { color: C.sub })), { html: span(when, { color: C.sub, size: 10 }), align: 'right' }]], { widths: ['12.4cm', '3.6cm'], borders: 'none' });
  if (items) await d.list(items);
};
/** The closing of a Chinese letter (此致 / 敬礼) or the English sign-off, then the sender and the date on the right. */
const signOff = async (d, t, who, date) => {
  if (t(true, false)) { await d.p(''); await d.p(esc('此致')); await d.p(esc('敬礼！')); }
  else { await d.p(''); await d.p(esc('Sincerely,')); await d.p(''); }
  await d.p(esc(who), { align: t('right', 'left') });
  await d.p(esc(date), { align: t('right', 'left') });
};
const check = (...items) => items.map(x => '☐ ' + x).join('　');

export default [
  {
    id: 'resume-classic', cat: '简历求职', name: ['简历 · 经典', 'Résumé · Classic'],
    async build(d, t) {
      const C = NAVY;
      await d.page({ page: 'A4', margin: 'moderate' });
      await d.p(span(t('李明', 'Alex Chen'), { size: 24, bold: true, color: C.acc }));
      await d.p(span(t('求职意向：产品经理', 'Product Manager'), { size: 12, color: C.acc }));
      await d.p(span(t('138 0000 0000　·　liming@example.com　·　上海　·　3 年产品经验', '+1 (555) 010-0000 · alex.chen@example.com · Seattle, WA · 3 years in product'), { size: 10, color: C.sub }));
      await d.rule(C.acc, 1.5);
      await section(d, C, t('工作经历', 'Experience'));
      await entry(d, C, t('远山科技有限公司', 'Brightline Technologies'), t('产品经理 · 企业协作产品线', 'Product Manager · Collaboration suite'), t('2023.07 – 至今', 'Jul 2023 – present'), [
        t('负责文档协作模块从 0 到 1，上线 6 个月月活达到 <b>12 万</b>，次月留存 41%。', 'Owned the document collaboration module from zero to launch: <b>120k</b> monthly actives in six months, 41% second-month retention.'),
        t('主导 3 次版本规划，协调设计、研发 15 人团队，按时交付率 95%。', 'Ran three release cycles with a 15-person design and engineering team; 95% of milestones delivered on time.'),
        t('建立用户反馈闭环，将高频需求的响应周期从 6 周缩短到 2 周。', 'Built the feedback loop that cut the turnaround on top requests from six weeks to two.'),
      ]);
      await entry(d, C, t('明川文化传媒', 'Riverstone Media'), t('产品助理', 'Associate Product Manager'), t('2022.07 – 2023.06', 'Jul 2022 – Jun 2023'), [
        t('负责内容后台的需求分析与原型设计，编辑效率提升 30%。', 'Wrote requirements and prototypes for the content back office; editor throughput up 30%.'),
        t('搭建数据看板，支撑内容团队每周复盘。', 'Set up the dashboards the content team reviews every week.'),
      ]);
      await section(d, C, t('项目经历', 'Projects'));
      await entry(d, C, t('智能会议纪要', 'Smart Meeting Notes'), t('项目负责人', 'Project lead'), '2024.03 – 2024.09', [
        t('把语音转写与要点提炼接入协作文档，覆盖公司 80% 的例会场景。', 'Brought transcription and summaries into the shared documents; now used in 80% of recurring meetings.'),
        t('上线后单场会议整理时间从 40 分钟降至 5 分钟。', 'Cut the time to write up a meeting from 40 minutes to five.'),
      ]);
      await section(d, C, t('教育背景', 'Education'));
      await entry(d, C, t('江城大学', 'Lakeside University'), t('工商管理 · 本科', 'B.A. in Business Administration'), t('2018.09 – 2022.06', 'Sep 2018 – Jun 2022'));
      await section(d, C, t('技能与证书', 'Skills'));
      await d.list([
        t('产品：需求分析、原型设计（Figma、Axure）、数据分析（SQL、Excel）', 'Product: discovery, prototyping (Figma, Axure), analysis (SQL, spreadsheets)'),
        t('证书：PMP、大学英语六级', 'Certificates: PMP; working proficiency in Mandarin'),
      ]);
      await section(d, C, t('自我评价', 'Summary'));
      await d.p(t('关注用户真实场景，习惯用数据说话；擅长在多方目标中找到可落地的方案，并推动团队按节奏交付。', 'Grounded in how people actually work and in the numbers; good at turning competing goals into a plan the team can ship on schedule.'));
    },
  },
  {
    id: 'resume-modern', cat: '简历求职', name: ['简历 · 双栏', 'Résumé · Two Columns'],
    async build(d, t) {
      const C = PLUM, W = 'FFFFFF';
      await d.page({ page: 'A4', margin: '1.4cm' });
      const tbl = await d.table([[{ html: span(t('王芳', 'Jordan Lee'), { size: 22, bold: true, color: W }), fill: C.acc, valign: 'top' }, { html: span(t('工作经历', 'Experience'), { size: 14, bold: true, color: C.acc }), valign: 'top' }]], { widths: ['6cm', '12.2cm'], borders: 'none', heights: [24.6] });
      const L = d.into(`${tbl}/row[1]/cell[1]`), R = d.into(`${tbl}/row[1]/cell[2]`);
      const side = async (title, items) => { await L.p(''); await L.p(span(title, { size: 11, bold: true, color: W })); for (const i of items) await L.p(span(i, { size: 10, color: 'E7E1F0' })); };
      await L.p(span(t('视觉设计师', 'Visual Designer'), { size: 11, color: 'E7E1F0' }));
      await side(t('联系方式', 'Contact'), [t('139 0000 0000', '+1 (555) 010-0000'), 'jordan.lee@example.com', t('杭州', 'Portland, OR'), 'jordanlee.design']);
      await side(t('技能', 'Skills'), ['Figma / Sketch', 'Photoshop / Illustrator', 'After Effects', t('设计系统与组件库', 'Design systems'), t('用户研究基础', 'User research basics')]);
      await side(t('语言', 'Languages'), [t('普通话（母语）', 'English (native)'), t('英语（CET-6，流利）', 'Spanish (conversational)')]);
      await side(t('荣誉', 'Awards'), [t('2025 年度公司设计奖', '2025 Company Design Award'), t('大学生广告艺术大赛一等奖', 'Student Advertising Awards, first prize')]);
      const job = async (org, role, when, items) => {
        await R.p(lines(bold(org), span(role + '　' + when, { color: C.sub, size: 10 })));
        for (const i of items) await R.p(esc(i), { list: 'bullet' });
      };
      await R.rule(C.line);
      await job(t('远山科技有限公司', 'Brightline Technologies'), t('视觉设计师', 'Visual Designer'), t('2023.03 – 至今', 'Mar 2023 – present'), [
        t('负责品牌视觉升级与全线产品的界面规范，落地 120 个组件的设计系统。', 'Led the brand refresh and the interface guidelines across the product line; a 120-component design system in use by four teams.'),
        t('主导 3 场发布会的视觉设计，物料交付零返工。', 'Designed the visuals for three launch events, delivered without a single rework round.'),
      ]);
      await job(t('明川文化传媒', 'Riverstone Media'), t('平面设计师', 'Graphic Designer'), t('2021.07 – 2023.02', 'Jul 2021 – Feb 2023'), [
        t('完成 40 余个营销项目的视觉设计，客户复购率 70%。', 'Delivered visuals for 40+ marketing campaigns; 70% of clients came back.'),
        t('搭建团队素材库与模板，制作效率提升 35%。', 'Built the shared asset library and templates that made production 35% faster.'),
      ]);
      await R.p(''); await R.p(span(t('教育背景', 'Education'), { size: 14, bold: true, color: C.acc })); await R.rule(C.line);
      await R.p(lines(bold(t('江城大学 · 视觉传达设计 · 本科', 'Lakeside University · B.F.A. in Visual Communication')), span(t('2017.09 – 2021.06', 'Sep 2017 – Jun 2021'), { color: C.sub, size: 10 })));
      await R.p(''); await R.p(span(t('代表项目', 'Selected work'), { size: 14, bold: true, color: C.acc })); await R.rule(C.line);
      await R.p(lines(bold(t('「协作」品牌重塑', '“Together” brand refresh')), esc(t('从标识到官网、应用图标与线下物料的完整视觉体系，上线后品牌认知度提升 18%。', 'Identity, website, app icons and print, as one system; brand recall up 18% after launch.'))));
      await R.p(lines(bold(t('年度设计报告', 'Annual design report')), esc(t('用数据可视化呈现全年产品迭代，成为公司对外的品牌资料。', 'The year in product, told through data graphics; now part of the company’s press kit.'))));
      await R.p(''); await R.p(span(t('自我评价', 'About me'), { size: 14, bold: true, color: C.acc })); await R.rule(C.line);
      await R.p(t('相信好的设计是把复杂留给自己、把简单留给用户。重视协作与反馈，能在节奏紧张的项目里稳定交付。', 'Good design keeps the complexity on our side and gives people the simple version. I work in the open, take feedback well and deliver steadily when the schedule is tight.'));
    },
  },
  {
    id: 'cover-letter', cat: '简历求职', name: ['求职信', 'Cover Letter'],
    async build(d, t) {
      const C = SLATE;
      await d.page({ page: 'A4', margin: 'normal' });
      await d.p(span(t('李明', 'Alex Chen'), { size: 16, bold: true, color: C.acc }));
      await d.p(span(t('138 0000 0000　·　liming@example.com　·　上海', '+1 (555) 010-0000 · alex.chen@example.com · Seattle, WA'), { size: 10, color: C.sub }));
      await d.rule(C.acc, 1.5);
      await d.p('');
      await d.p(esc(t('2026 年 9 月 25 日', 'September 25, 2026')));
      await d.p(lines(esc(t('远山科技有限公司 人力资源部', 'Hiring Team')), esc(t('产品经理岗位招聘负责人', 'Brightline Technologies')), t(null, esc('410 Pine Street, Seattle, WA 98101'))));
      await d.p('');
      await d.p(esc(t('尊敬的招聘负责人：', 'Dear Hiring Team,')));
      await d.p(t('我从贵公司官网看到产品经理岗位的招聘信息，结合我三年 B 端协作产品的经验，希望申请这一职位。', 'I am writing to apply for the Product Manager position posted on your careers page. Three years of building collaboration products for business teams have prepared me well for it.'));
      await d.p(t('在远山科技，我负责文档协作模块从 0 到 1 的规划与交付：上线六个月月活用户达到 12 万，次月留存 41%。这个过程让我学会在多方目标之间找到可落地的方案，并带领 15 人的设计与研发团队按节奏交付。', 'At my current company I took the document collaboration module from an idea to launch: 120,000 monthly active users within six months and 41% second-month retention. Along the way I learned to turn competing goals into a plan a 15-person design and engineering team could deliver on schedule.'));
      await d.p(t('我关注贵公司在企业协作领域的产品思路，尤其认同「让工具退到工作背后」的理念。我在用户反馈闭环和数据驱动决策上的积累，能帮助团队更快地验证方向、减少返工。', 'What draws me to your team is the way your products get out of the way of the work. My experience with feedback loops and data-driven decisions would help the team validate directions sooner and rework less.'));
      await d.p(t('随信附上我的简历，期待有机会进一步交流。感谢您的时间。', 'My résumé is attached. I would welcome the chance to talk further, and I appreciate your time.'));
      await signOff(d, t, t('李明', 'Alex Chen'), t('2026 年 9 月 25 日', ''));
    },
  },
  {
    id: 'work-report', cat: '报告计划', name: ['工作报告', 'Work Report'],
    async build(d, t) {
      const C = NAVY;
      await d.page({ page: 'A4', margin: 'normal', header: span(t('远山科技有限公司 · 内部文件', 'Brightline Technologies · Internal'), { color: C.sub, size: 9 }), footer: t('第 {page} 页 / 共 {pages} 页', 'Page {page} of {pages}') });
      await d.title(span(t('2026 年上半年工作报告', 'Half-Year Work Report, 2026'), { color: C.acc }));
      await d.p(span(t('市场部　·　汇报人：张伟　·　2026 年 7 月 10 日', 'Marketing · Prepared by Sam Taylor · July 10, 2026'), { color: C.sub, size: 10 }));
      await heading(d, C, t('一、工作概述', '1. Overview'), 1);
      await d.p(t('上半年市场部围绕「品牌升级、渠道拓展、会员增长」三条主线开展工作。截至 6 月 30 日，年度目标整体完成 54%，其中会员增长提前达到进度，品牌曝光略低于计划。', 'In the first half the team worked along three lines: the brand refresh, new channels and membership growth. By June 30 we had delivered 54% of the annual targets; membership is ahead of plan, brand reach slightly behind.'));
      await heading(d, C, t('二、主要工作完成情况', '2. Results'), 1);
      await dataTable(d, C, [t('工作事项', 'Item'), t('年度目标', 'Annual target'), t('上半年完成', 'Delivered'), t('完成率', 'Progress')], [
        [t('品牌曝光（万次）', 'Brand reach (millions)'), '12,000', '5,600', '47%'],
        [t('新增会员（人）', 'New members'), '80,000', '46,300', '58%'],
        [t('渠道合作（家）', 'Channel partners'), '10', '6', '60%'],
        [t('市场活动（场）', 'Events'), '24', '13', '54%'],
      ], { widths: ['6cm', '3.4cm', '3.4cm', '3.2cm'] });
      await d.p('');
      await d.list([
        t('<b>品牌升级</b>：完成新视觉体系并在官网、门店和主要渠道落地，品牌认知度调研提升 11 个百分点。', '<b>Brand refresh</b>: the new visual system is live on the website, in stores and across the main channels; recall up 11 points in the survey.'),
        t('<b>渠道拓展</b>：新签 6 家区域渠道，华南、西南两区实现零的突破。', '<b>Channels</b>: six regional partners signed, including the first in the South and Southwest.'),
        t('<b>会员增长</b>：会员日活动与积分体系上线后，月均新增会员由 5,800 提升至 9,200。', '<b>Membership</b>: monthly sign-ups rose from 5,800 to 9,200 after the members’ day and the points programme launched.'),
      ]);
      await heading(d, C, t('三、存在的问题', '3. Issues'), 1);
      await d.list([
        t('品牌曝光低于计划，主要因二季度投放预算延后到位，线上投放晚启动 5 周。', 'Brand reach fell short of plan: the Q2 media budget arrived late and online placements started five weeks behind.'),
        t('部分区域渠道的培训与物料支持不足，首月动销低于预期。', 'Some regional partners lacked training and materials, and their first-month sell-through was below expectations.'),
        t('跨部门项目排期冲突较多，需求响应周期偏长。', 'Scheduling conflicts across departments stretched response times.'),
      ]);
      await heading(d, C, t('四、下半年工作计划', '4. Plan for the second half'), 1);
      await d.list([
        t('集中资源完成秋季新品上市与年终大促两场关键战役，补足曝光缺口。', 'Concentrate on the two key campaigns, the autumn launch and the year-end sale, to close the reach gap.'),
        t('建立渠道伙伴的标准化培训包与月度巡访机制。', 'Standardise the partner training kit and start monthly partner visits.'),
        t('会员体系二期上线，目标全年新增会员 9 万。', 'Launch phase two of the membership programme, aiming at 90,000 new members for the year.'),
        t('与产品、IT 部门建立双周需求评审，缩短响应周期。', 'Set up a fortnightly review with Product and IT to shorten response times.'),
      ], { list: 'number' });
      await heading(d, C, t('五、需要的支持', '5. Support needed'), 1);
      await d.p(t('建议三季度追加渠道支持预算 30 万元，并由人力资源部协助补充 2 名区域市场专员。', 'A further 300,000 in channel support for Q3, and two regional marketing specialists recruited with HR’s help.'));
    },
  },
  {
    id: 'weekly-report', cat: '报告计划', name: ['周报', 'Weekly Report'],
    async build(d, t) {
      const C = NAVY;
      await d.page({ page: 'A4', margin: 'normal', footer: t('第 {page} 页 / 共 {pages} 页', 'Page {page} of {pages}') });
      await d.title(span(t('工作周报', 'Weekly Report'), { color: C.acc }));
      await d.p(span(t('2026 年第 39 周　9 月 21 日—9 月 25 日', 'Week 39, 2026 · September 21–25'), { color: C.sub, size: 10 }));
      await infoTable(d, C, [[t('姓名', 'Name'), t('林晓', 'Alex Chen')], [t('部门', 'Team'), t('市场部', 'Marketing')], [t('岗位', 'Role'), t('品牌经理', 'Brand Manager')], [t('汇报对象', 'Reports to'), t('市场总监', 'Head of Marketing')]]);
      await heading(d, C, t('一、本周完成', '1. Done this week'));
      await dataTable(d, C, [t('工作事项', 'Item'), t('进度', 'Status'), t('说明', 'Notes')], [
        [t('秋季新品预热方案', 'Autumn launch teaser plan'), t('已完成', 'Done'), t('方案已定稿，门店物料下发', 'Plan signed off; store materials sent out')],
        [t('社交平台内容排期', 'Social content calendar'), '80%', t('10 月内容已排期，待法务审核', 'October posts scheduled, awaiting legal review')],
        [t('会员活动数据复盘', 'Member campaign review'), t('已完成', 'Done'), t('复购率 32%，高于目标 2 个百分点', 'Repeat rate 32%, 2 points above target')],
        [t('渠道合作洽谈', 'Channel partnerships'), '50%', t('两家渠道进入报价阶段', 'Two partners are at the quote stage')],
      ], { widths: ['5cm', '2.4cm', '8.6cm'] });
      await heading(d, C, t('二、关键数据', '2. Key numbers'));
      await d.list([
        t('内容曝光 <b>536 万</b>次，完成周目标的 107%。', 'Content reach <b>5.36M</b>, 107% of the weekly target.'),
        t('新增会员 <b>4,280</b> 人，环比增长 12%。', 'New members <b>4,280</b>, up 12% week over week.'),
        t('活动投入 3.2 万元，单个新会员成本 7.5 元。', 'Campaign spend $4,400; cost per new member $1.03.'),
      ]);
      await heading(d, C, t('三、下周计划', '3. Next week'));
      await d.list([
        t('新品首发活动执行与门店巡检。', 'Run the launch event and visit the stores.'),
        t('完成 10 月内容的法务审核并上线。', 'Clear legal review for October content and publish it.'),
        t('与两家渠道确定合作条款。', 'Agree terms with the two channel partners.'),
      ], { list: 'number' });
      await heading(d, C, t('四、需要协调', '4. Help needed'));
      await d.p(t('华南区首批物料晚到两天，需要供应链确认补发时间；首发活动需要门店运营部安排 2 名支援人员。', 'The first batch of South-region materials arrived two days late: supply chain to confirm the reshipment date. Store operations to lend two people for the launch event.'));
    },
  },
  {
    id: 'meeting-minutes', cat: '行政办公', name: ['会议纪要', 'Meeting Minutes'],
    async build(d, t) {
      const C = INK;
      await d.page({ page: 'A4', margin: 'normal', footer: t('第 {page} 页 / 共 {pages} 页', 'Page {page} of {pages}') });
      await d.title(span(t('会议纪要', 'Meeting Minutes'), { color: C.acc }));
      await d.p(span(t('秋季新品上市项目 · 第 4 次例会', 'Autumn Launch Project · Meeting 4'), { color: C.sub, size: 10 }));
      await d.table([
        [{ html: span(t('会议时间', 'Date'), { bold: true }), fill: C.soft }, esc(t('2026 年 9 月 24 日 14:00–15:30', 'September 24, 2026, 2:00–3:30 pm')), { html: span(t('地点', 'Place'), { bold: true }), fill: C.soft }, esc(t('3 楼 302 会议室', 'Room 302')) ],
        [{ html: span(t('主持人', 'Chair'), { bold: true }), fill: C.soft }, esc(t('张伟', 'Sam Taylor')), { html: span(t('记录人', 'Minutes'), { bold: true }), fill: C.soft }, esc(t('陈静', 'Morgan Reed'))],
        [{ html: span(t('参会人员', 'Attendees'), { bold: true }), fill: C.soft }, { html: esc(t('市场部：张伟、林晓、陈静；产品部：李明；供应链：刘洋；门店运营：王芳', 'Marketing: Sam Taylor, Alex Chen, Morgan Reed · Product: Jordan Lee · Supply chain: Casey Kim · Stores: Riley Park')), colspan: 3 }, null, null],
      ], { widths: ['2.6cm', '5.4cm', '2.6cm', '5.4cm'], borders: 'all', borderColor: C.line });
      await heading(d, C, t('一、议题与结论', '1. Topics and decisions'));
      await dataTable(d, C, [t('议题', 'Topic'), t('讨论要点', 'Discussion'), t('结论', 'Decision')], [
        [t('首发活动方案', 'Launch event plan'), t('两套方案的成本与覆盖对比；门店侧倾向方案 B', 'Cost and reach of the two options; stores prefer option B'), t('采用方案 B，预算上限 45 万元', 'Option B, budget capped at 450,000')],
        [t('物料到货进度', 'Materials delivery'), t('华南区首批物料延迟两天，其余区域正常', 'The South batch is two days late, other regions on time'), t('供应链 9 月 26 日前确认补发', 'Supply chain to confirm reshipment by Sep 26')],
        [t('线上内容排期', 'Online content'), t('10 月内容待法务审核，预计 3 个工作日', 'October content awaits legal review, about three working days'), t('9 月 30 日前上线首批内容', 'First batch live by Sep 30')],
      ], { widths: ['3.4cm', '7.2cm', '5.4cm'] });
      await heading(d, C, t('二、待办事项', '2. Action items'));
      await dataTable(d, C, [t('事项', 'Action'), t('负责人', 'Owner'), t('截止日期', 'Due'), t('状态', 'Status')], [
        [t('确认华南区物料补发时间', 'Confirm the South reshipment date'), t('刘洋', 'Casey Kim'), '09-26', t('进行中', 'In progress')],
        [t('提交方案 B 的执行细案', 'Submit the detailed plan for option B'), t('林晓', 'Alex Chen'), '09-29', t('未开始', 'Not started')],
        [t('跟进法务审核并安排上线', 'Follow up legal review and schedule publishing'), t('陈静', 'Morgan Reed'), '09-30', t('进行中', 'In progress')],
        [t('门店支援人员排班', 'Roster the store support staff'), t('王芳', 'Riley Park'), '10-08', t('未开始', 'Not started')],
      ], { widths: ['7cm', '2.6cm', '2.8cm', '3.6cm'] });
      await heading(d, C, t('三、其他', '3. Other'));
      await d.p(t('下次例会定于 10 月 8 日 14:00，同一地点；请各负责人会前更新待办状态。', 'The next meeting is on October 8 at 2:00 pm, same room. Owners please update their action items beforehand.'));
    },
  },
  {
    id: 'notice', cat: '行政办公', name: ['通知', 'Notice'],
    async build(d, t) {
      const RED = 'C00000';
      await d.page({ page: 'A4', margin: 'normal' });
      if (t(true, false)) {
        await d.p(span('远山科技有限公司文件', { size: 26, bold: true, color: RED }), { align: 'center' });
        await d.p(span('远科发〔2026〕12 号', { size: 12 }), { align: 'center' });
        await d.rule(RED, 2);
        await d.p('');
        await d.p(span('关于开展 2026 年度员工健康体检工作的通知', { size: 16, bold: true }), { align: 'center' });
        await d.p('');
        await d.p('各部门、各分公司：');
        await d.p('为保障员工身体健康，公司决定开展 2026 年度员工健康体检工作。现将有关事项通知如下：');
        await d.h3('一、体检对象');
        await d.p('截至 2026 年 9 月 30 日已转正的全体员工。');
        await d.h3('二、体检时间与地点');
        await d.p('2026 年 10 月 13 日至 10 月 24 日，工作日上午 8:00—11:00，请空腹前往。地点：市中心医院体检中心（人民路 88 号）。');
        await d.h3('三、体检安排');
        await d.list(['各部门按附件分批次参加体检，特殊情况可与行政部另行约定时间。', '体检当日凭工牌与身份证签到，体检项目见附件二。', '体检费用由公司统一承担，员工可自费加项。'], { list: 'number' });
        await d.h3('四、其他事项');
        await d.p('请各部门负责人做好组织工作，确保员工按时参加。未尽事宜，请联系行政部（分机 8021）。');
        await d.p('特此通知。');
        await d.p('');
        await d.p('附件：1. 各部门体检批次安排表　2. 体检项目清单');
        await d.p('');
        await d.p('远山科技有限公司行政部', { align: 'right' });
        await d.p('2026 年 9 月 25 日', { align: 'right' });
      } else {
        const C = NAVY;
        await d.p(span('MEMORANDUM', { size: 24, bold: true, color: C.acc }));
        await d.p(span('Brightline Technologies · Human Resources', { size: 10, color: C.sub }));
        await d.rule(C.acc, 1.5);
        await d.table([
          [{ html: bold('To'), fill: C.soft }, 'All employees'], [{ html: bold('From'), fill: C.soft }, 'Human Resources'],
          [{ html: bold('Date'), fill: C.soft }, 'September 25, 2026'], [{ html: bold('Subject'), fill: C.soft }, bold('2026 annual health check-ups')],
        ], { widths: ['3cm', '13cm'], borders: 'horizontal', borderColor: C.line });
        await d.p('');
        await d.p('The company is arranging this year’s health check-ups for all employees who have completed their probation by September 30, 2026. Please note the following.');
        await d.h3('When and where');
        await d.p('October 13–24, 2026, on working days between 8:00 and 11:00 am, at the City Central Hospital screening centre (88 Renmin Road). Please come fasting.');
        await d.h3('How it works');
        await d.list(['Departments attend in the batches listed in attachment 1; Administration can arrange another date in special cases.', 'Sign in with your staff badge and ID; the examination list is in attachment 2.', 'The company covers the standard package; extra items can be added at your own expense.'], { list: 'number' });
        await d.h3('Questions');
        await d.p('Department heads, please make sure your teams attend on time. For anything else, contact Administration at extension 8021.');
        await d.p('');
        await d.p('Attachments: 1. Check-up schedule by department  2. Examination list');
      }
    },
  },
  {
    id: 'contract', cat: '合同信函', name: ['合同', 'Contract'],
    async build(d, t) {
      const C = INK;
      await d.page({ page: 'A4', margin: 'normal', footer: t('第 {page} 页 / 共 {pages} 页', 'Page {page} of {pages}') });
      await d.p(span(t('技术服务合同', 'Service Agreement'), { size: 22, bold: true }), { align: 'center' });
      await d.p(span(t('合同编号：YS-2026-0925', 'Agreement No. BL-2026-0925'), { color: C.sub, size: 10 }), { align: 'right' });
      await d.table([
        [{ html: bold(t('甲方（委托方）', 'Client')), fill: C.soft }, esc(t('明川文化传媒有限公司', 'Riverstone Media Ltd.')), { html: bold(t('乙方（受托方）', 'Provider')), fill: C.soft }, esc(t('远山科技有限公司', 'Brightline Technologies Ltd.'))],
        [{ html: bold(t('地址', 'Address')), fill: C.soft }, esc(t('上海市浦东新区世纪大道 100 号', '100 Century Avenue, Shanghai')), { html: bold(t('地址', 'Address')), fill: C.soft }, esc(t('杭州市西湖区文三路 258 号', '258 Wensan Road, Hangzhou'))],
        [{ html: bold(t('联系人', 'Contact')), fill: C.soft }, esc(t('张伟　138 0000 0001', 'Sam Taylor · +86 138 0000 0001')), { html: bold(t('联系人', 'Contact')), fill: C.soft }, esc(t('李明　139 0000 0002', 'Alex Chen · +86 139 0000 0002'))],
      ], { widths: ['2.8cm', '5.2cm', '2.8cm', '5.2cm'], borders: 'all', borderColor: C.line });
      await d.p('');
      await d.p(t('鉴于甲方需要委托乙方提供软件开发与技术服务，双方本着平等自愿、诚实信用的原则，经友好协商，订立本合同，共同遵守。', 'The Client wishes to engage the Provider for software development and technical services. In good faith and on equal terms, the parties agree as follows.'));
      const clause = async (title, paras) => { await d.h2(esc(title)); for (const p of paras) await d.p(p); };
      await clause(t('第一条　服务内容', '1. Services'), [t('乙方为甲方提供内容管理系统的定制开发、部署及为期 12 个月的运维支持，具体范围以附件一《需求说明书》为准。', 'The Provider will design, build and deploy a content management system for the Client and support it for twelve months, as described in Schedule 1 (Statement of Work).')]);
      await clause(t('第二条　合同期限', '2. Term'), [t('本合同自 2026 年 10 月 1 日起至 2027 年 9 月 30 日止。期满前 30 日内，双方可协商续签。', 'This agreement runs from October 1, 2026 to September 30, 2027. The parties may agree to renew it within 30 days before it ends.')]);
      await clause(t('第三条　合同金额与支付', '3. Fees and payment'), [
        t('合同总金额为人民币 480,000 元（大写：肆拾捌万元整），含税。', 'The total fee is 480,000, inclusive of tax.'),
        t('付款方式：合同签订后 10 个工作日内支付 30%；系统验收合格后 10 个工作日内支付 60%；运维期满后 10 个工作日内支付剩余 10%。', 'Payment: 30% within ten working days of signing; 60% within ten working days of acceptance; the remaining 10% within ten working days after the support period ends.'),
      ]);
      await clause(t('第四条　双方权利义务', '4. Obligations'), [
        t('甲方应按约定提供必要的资料与配合，并按期支付款项。', 'The Client will provide the materials and cooperation reasonably required and pay on time.'),
        t('乙方应按附件约定的进度与质量交付成果，并对交付物提供保修。', 'The Provider will deliver to the schedule and quality set out in the Schedules and warrant the deliverables.'),
      ]);
      await clause(t('第五条　知识产权与保密', '5. Intellectual property and confidentiality'), [t('为本合同开发的定制成果，其知识产权在甲方付清全部款项后归甲方所有；乙方原有的通用组件仍归乙方所有并授权甲方使用。双方对在履行本合同过程中知悉的对方商业秘密负有保密义务，保密期限不受本合同期限限制。', 'Custom work created under this agreement belongs to the Client once all fees are paid; the Provider keeps its pre-existing components and licenses the Client to use them. Each party will keep the other’s confidential information secret, and this duty survives the end of the agreement.')]);
      await clause(t('第六条　违约责任', '6. Liability'), [t('任何一方违反本合同约定，应承担由此给对方造成的直接损失。乙方逾期交付的，每逾期一日按合同总额的 0.1% 支付违约金，累计不超过合同总额的 10%。', 'A party in breach will compensate the other for the direct loss it causes. For late delivery the Provider pays 0.1% of the total fee per day, capped at 10% of the total fee.')]);
      await clause(t('第七条　争议解决与其他', '7. Disputes and general'), [t('因本合同引起的争议，双方应友好协商解决；协商不成的，提交甲方所在地有管辖权的人民法院诉讼解决。本合同一式两份，双方各执一份，自双方签字盖章之日起生效。', 'The parties will try to settle any dispute amicably; failing that, the courts where the Client is located have jurisdiction. This agreement is signed in two counterparts, one for each party, and takes effect when both have signed.')]);
      await d.p('');
      await d.table([
        [{ html: lines(bold(t('甲方（盖章）：', 'Client')), '', esc(t('授权代表（签字）：', 'Authorised signature:')), '', esc(t('日期：　　年　　月　　日', 'Name and title:')), t(null, ''), t(null, esc('Date:'))), valign: 'top' },
         { html: lines(bold(t('乙方（盖章）：', 'Provider')), '', esc(t('授权代表（签字）：', 'Authorised signature:')), '', esc(t('日期：　　年　　月　　日', 'Name and title:')), t(null, ''), t(null, esc('Date:'))), valign: 'top' }],
      ], { widths: ['8cm', '8cm'], borders: 'none', heights: [4] });
    },
  },
  {
    id: 'letter', cat: '合同信函', name: ['商务信函', 'Business Letter'],
    async build(d, t) {
      const C = SLATE;
      await d.page({ page: 'A4', margin: 'normal' });
      await d.p(span(t('远山科技有限公司', 'Brightline Technologies'), { size: 16, bold: true, color: C.acc }));
      await d.p(span(t('杭州市西湖区文三路 258 号　·　0571-0000 0000　·　hello@example.com', '258 Wensan Road, Hangzhou · +86 571 0000 0000 · hello@example.com'), { size: 9.5, color: C.sub }));
      await d.rule(C.acc, 1.5);
      await d.p('');
      await d.p(esc(t('2026 年 9 月 25 日', 'September 25, 2026')));
      await d.p('');
      await d.p(lines(esc(t('明川文化传媒有限公司', 'Ms. Morgan Reed')), esc(t('张伟 总经理', 'General Manager, Riverstone Media Ltd.')), esc(t('上海市浦东新区世纪大道 100 号', '100 Century Avenue, Shanghai'))));
      await d.p('');
      await d.p(esc(t('尊敬的张总：', 'Dear Ms. Reed,')));
      await d.p(t('感谢贵公司对我们内容管理系统的关注，也感谢您上周抽出时间参加产品演示。现就双方讨论的合作事项，向您说明我们的建议方案。', 'Thank you for your interest in our content management system, and for taking the time to join last week’s demonstration. This letter sets out the proposal we discussed.'));
      await d.p(t('根据贵司的需求，我们建议分两期实施：第一期（10 月—12 月）完成核心内容模块与权限体系的部署，第二期（次年 1 月—3 月）接入数据分析与多渠道发布。整体报价为人民币 48 万元，含 12 个月的运维支持，详细说明见随信附件。', 'We suggest two phases: the core content modules and permissions from October to December, then analytics and multi-channel publishing from January to March. The total comes to 480,000 including twelve months of support; the attached proposal has the details.'));
      await d.p(t('如方案符合贵司预期，我们可于 10 月上旬安排启动会议。如有任何疑问，欢迎随时与我联系。', 'If this matches your expectations, we can hold the kick-off meeting in early October. Please contact me with any questions.'));
      await d.p(t('期待与贵公司的合作。', 'We look forward to working with you.'));
      await signOff(d, t, t('远山科技有限公司　李明', 'Alex Chen'), t('客户总监', 'Director of Client Services'));
      await d.p('');
      await d.p(span(t('附件：《内容管理系统实施方案》', 'Enclosure: Implementation proposal'), { color: C.sub, size: 10 }));
    },
  },
  {
    id: 'thesis', cat: '学习', name: ['论文封面与目录', 'Thesis Cover & Contents'],
    async build(d, t) {
      const C = INK;
      await d.page({ page: 'A4', margin: 'normal', header: span(t('基于用户行为数据的协作工具留存分析', 'Retention in Collaboration Tools: Evidence from Behavioural Data'), { color: C.sub, size: 9 }), footer: t('第 {page} 页', 'Page {page}'), titlePg: true });
      await d.p(''); await d.p('');
      await d.p(span(t('江城大学', 'Lakeside University'), { size: 26, bold: true }), { align: 'center' });
      await d.p(span(t('本科毕业论文', 'Undergraduate Thesis'), { size: 16, color: C.sub }), { align: 'center' });
      await d.p(''); await d.p(''); await d.p('');
      await d.p(span(t('基于用户行为数据的协作工具留存分析', 'Retention in Collaboration Tools: Evidence from Behavioural Data'), { size: 20, bold: true }), { align: 'center' });
      await d.p(''); await d.p(''); await d.p('');
      const pairs = [[t('学　　院', 'School'), t('管理学院', 'School of Management')], [t('专　　业', 'Programme'), t('工商管理', 'Business Administration')], [t('学　　号', 'Student ID'), '2022101234'], [t('姓　　名', 'Author'), t('李明', 'Alex Chen')], [t('指导教师', 'Supervisor'), t('王芳 教授', 'Prof. Jordan Lee')], [t('完成日期', 'Date'), t('2026 年 5 月', 'May 2026')]];
      await d.table(pairs.map(([k, v]) => [{ html: span(k, { size: 13 }), align: 'right', borders: 'none' }, { html: span(v, { size: 13 }), align: 'center' }]), { widths: ['4cm', '7cm'], borders: 'horizontal', borderColor: C.acc, align: 'center', heights: pairs.map(() => 0.95) });
      await d.pagebreak();
      await d.toc({ title: t('目　录', 'Contents'), levels: 2 });
      await d.pagebreak();
      await d.h1(esc(t('摘　要', 'Abstract')));
      await d.p(t('本文以某协作工具 2025 年的用户行为日志为样本，考察功能使用深度、协作规模与留存之间的关系。研究发现，首周内完成一次多人协作的用户，次月留存率显著高于仅单人使用者；功能使用广度对留存的影响呈边际递减。据此提出面向新用户引导的三点建议。', 'Using the 2025 behavioural logs of a collaboration tool, this thesis examines how depth of feature use and the size of collaboration relate to retention. Users who complete one multi-person collaboration in their first week retain markedly better in the second month than solo users, and the effect of feature breadth diminishes at the margin. Three recommendations for new-user onboarding follow.'));
      await d.p(bold(t('关键词：', 'Keywords: ')) + esc(t('用户留存；协作工具；行为数据；新用户引导', 'user retention; collaboration tools; behavioural data; onboarding')));
      await d.h1(esc(t('第一章　绪论', 'Chapter 1  Introduction')));
      await d.h2(esc(t('1.1　研究背景', '1.1  Background')));
      await d.p(t('协作工具已成为知识工作的基础设施，但多数产品在获取用户后面临明显的流失。理解哪些早期行为预示着长期留存，对产品设计具有直接意义。', 'Collaboration tools have become the infrastructure of knowledge work, yet most lose a large share of the users they acquire. Knowing which early behaviours predict long-term retention bears directly on product design.'));
      await d.h2(esc(t('1.2　研究问题与意义', '1.2  Research questions')));
      await d.p(t('本文试图回答：（1）哪些首周行为与次月留存显著相关；（2）协作规模对留存的作用是否存在阈值；（3）这些发现如何转化为新用户引导策略。', 'This thesis asks (1) which first-week behaviours correlate with second-month retention, (2) whether the effect of collaboration size has a threshold, and (3) how the findings translate into onboarding strategy.'));
      await d.h1(esc(t('第二章　文献综述', 'Chapter 2  Literature Review')));
      await d.h2(esc(t('2.1　用户留存的相关研究', '2.1  Studies of user retention')));
      await d.p(t('已有研究多从生命周期与漏斗视角刻画留存……', 'Prior work largely frames retention through lifecycle and funnel models…'));
      await d.h2(esc(t('2.2　协作行为与产品价值', '2.2  Collaboration and product value')));
      await d.p(t('网络效应理论认为，多人参与提高了单个用户的转换成本……', 'Network-effect theory holds that participation by others raises each user’s switching cost…'));
      await d.h1(esc(t('第三章　研究设计', 'Chapter 3  Method')));
      await d.h2(esc(t('3.1　数据来源', '3.1  Data')));
      await d.p(t('样本为 2025 年 1 月至 12 月注册的 48,000 名用户的脱敏行为日志……', 'The sample covers the anonymised logs of 48,000 users who registered between January and December 2025…'));
      await d.h2(esc(t('3.2　变量与模型', '3.2  Variables and model')));
      await d.p(t('因变量为次月是否活跃，自变量包括首周协作人数、功能使用广度与深度……', 'The dependent variable is second-month activity; the predictors include first-week collaborators, feature breadth and depth…'));
      await d.h1(esc(t('参考文献', 'References')));
      await d.list([
        t('王芳, 李明. 协作工具的用户留存研究[J]. 管理评论, 2025, 37(4): 112-120.', 'Lee, J., & Chen, A. (2025). User retention in collaboration tools. <i>Management Review</i>, 37(4), 112–120.'),
        t('Reed M. Network effects in productivity software[J]. Journal of Product Research, 2024, 12(2): 45-61.', 'Reed, M. (2024). Network effects in productivity software. <i>Journal of Product Research</i>, 12(2), 45–61.'),
      ], { list: 'number' });
      await d.set('/body/toc[1]', { levels: 2 });
    },
  },
  {
    id: 'proposal', cat: '报告计划', name: ['项目计划书', 'Project Plan'],
    async build(d, t) {
      const C = PLUM;
      await d.page({ page: 'A4', margin: 'normal', footer: t('第 {page} 页 / 共 {pages} 页', 'Page {page} of {pages}') });
      await d.p(span(t('项目计划书', 'Project Plan'), { size: 12, color: C.sub }));
      await d.p(span(t('新版官网改版项目', 'Website Redesign'), { size: 26, bold: true, color: C.acc }));
      await d.rule(C.acc, 1.5);
      await infoTable(d, C, [[t('编制部门', 'Owner'), t('市场部', 'Marketing')], [t('项目经理', 'Project manager'), t('林晓', 'Alex Chen')], [t('编制日期', 'Date'), t('2026 年 9 月 25 日', 'September 25, 2026')], [t('版本', 'Version'), 'v1.0']]);
      await heading(d, C, t('一、项目背景', '1. Background'), 1);
      await d.p(t('现有官网上线于 2021 年，信息架构以公司介绍为主，产品入口深、移动端体验差，近一年访问转化率持续下滑至 1.2%。随着新品牌视觉体系落地，官网需要同步改版，成为产品获客的主要入口。', 'The current website dates from 2021. It is built around the company profile, buries the products and works poorly on phones; conversion has slipped to 1.2% over the past year. With the new brand identity in place, the site needs a redesign that makes it the main acquisition channel.'));
      await heading(d, C, t('二、项目目标', '2. Goals'), 1);
      await d.list([
        t('11 月 30 日前完成新版官网上线，覆盖桌面端与移动端。', 'Launch the new site for desktop and mobile by November 30.'),
        t('上线三个月内访问转化率从 1.2% 提升到 2.0%。', 'Raise conversion from 1.2% to 2.0% within three months of launch.'),
        t('移动端页面加载时间控制在 2 秒以内，可用性评分 4.5 以上。', 'Mobile pages load in under two seconds; usability score of 4.5 or better.'),
      ]);
      await heading(d, C, t('三、项目范围', '3. Scope'), 1);
      await d.table([
        [{ html: span(t('范围内', 'In scope'), { bold: true, color: 'FFFFFF' }), fill: C.acc }, { html: span(t('范围外', 'Out of scope'), { bold: true, color: 'FFFFFF' }), fill: C.sub }],
        [lines(esc(t('· 信息架构与页面设计（约 24 个页面）', '· Information architecture and page design (about 24 pages)')), esc(t('· 内容迁移与新文案', '· Content migration and new copy')), esc(t('· 前端开发与内容管理系统接入', '· Front-end build and CMS integration')), esc(t('· 数据埋点与转化漏斗', '· Analytics and the conversion funnel'))),
         lines(esc(t('· 产品内文档站与帮助中心', '· The in-product docs and help centre')), esc(t('· 多语言版本（二期考虑）', '· Additional languages (phase two)')), esc(t('· 线上商城与支付', '· The online store and payments')))],
      ], { widths: ['8cm', '8cm'], borders: 'all', borderColor: C.line });
      await heading(d, C, t('四、实施计划', '4. Schedule'), 1);
      await dataTable(d, C, [t('阶段', 'Phase'), t('主要工作', 'Work'), t('起止时间', 'Dates'), t('负责人', 'Owner')], [
        [t('需求与架构', 'Discovery'), t('用户访谈、竞品分析、信息架构定稿', 'User interviews, competitor review, information architecture'), '09-28 – 10-10', t('林晓', 'Alex Chen')],
        [t('设计', 'Design'), t('视觉方案、页面设计、设计评审', 'Visual direction, page designs, design review'), '10-11 – 10-31', t('王芳', 'Jordan Lee')],
        [t('开发', 'Build'), t('前端开发、CMS 接入、内容迁移', 'Front end, CMS integration, content migration'), '10-20 – 11-15', t('刘洋', 'Casey Kim')],
        [t('测试与上线', 'Test & launch'), t('功能测试、性能优化、灰度发布', 'QA, performance tuning, staged rollout'), '11-16 – 11-30', t('陈静', 'Morgan Reed')],
      ], { widths: ['2.8cm', '7.2cm', '3.2cm', '2.8cm'] });
      await heading(d, C, t('五、资源与预算', '5. Budget'), 1);
      await dataTable(d, C, [t('预算项', 'Item'), t('金额（元）', 'Amount'), t('说明', 'Notes')], [
        [t('设计外包', 'Design contractor'), '120,000', t('插画与动效', 'Illustration and motion')],
        [t('开发人力', 'Engineering'), '150,000', t('3 人 × 2 个月', 'Three people for two months')],
        [t('内容与摄影', 'Content and photography'), '40,000', t('产品图与案例采访', 'Product photography and case studies')],
        [t('测试与工具', 'Testing and tools'), '20,000', t('可用性测试与监测服务', 'Usability testing and monitoring')],
        [{ html: bold(t('合计', 'Total')), fill: C.soft }, { html: bold('330,000'), fill: C.soft }, { html: '', fill: C.soft }],
      ], { widths: ['5cm', '3.6cm', '7.4cm'] });
      await heading(d, C, t('六、风险与应对', '6. Risks'), 1);
      await dataTable(d, C, [t('风险', 'Risk'), t('影响', 'Impact'), t('应对措施', 'Response')], [
        [t('内容迁移量大于预期', 'More content to migrate than expected'), t('高', 'High'), t('提前两周启动盘点，非核心页面二期迁移', 'Start the inventory two weeks early; move non-core pages in phase two')],
        [t('设计评审反复', 'Repeated design reviews'), t('中', 'Medium'), t('每阶段限定两轮评审，明确决策人', 'Two review rounds per phase with a named decision maker')],
        [t('第三方接口不稳定', 'Unreliable third-party APIs'), t('中', 'Medium'), t('关键接口准备降级方案', 'Fallbacks for the key integrations')],
      ], { widths: ['5cm', '2cm', '9cm'] });
      await heading(d, C, t('七、验收标准', '7. Acceptance'), 1);
      await d.list([t('全部页面通过功能测试，主流浏览器与手机机型显示正常。', 'Every page passes QA and renders correctly in the main browsers and on common phones.'), t('移动端性能评分 90 分以上。', 'Mobile performance score of 90 or higher.'), t('转化埋点数据完整，可在看板中查看。', 'Conversion events are complete and visible on the dashboard.')]);
    },
  },
  {
    id: 'leave-request', cat: '行政办公', name: ['请假申请单', 'Leave Request Form'],
    async build(d, t) {
      const C = SLATE, L = k => ({ html: span(k, { bold: true }), fill: C.soft, valign: 'middle' });
      await d.page({ page: 'A4', margin: 'normal' });
      await d.p(span(t('员工请假申请单', 'Leave Request Form'), { size: 20, bold: true, color: C.acc }), { align: 'center' });
      await d.p(span(t('远山科技有限公司 · 人力资源部　　单号：LR-2026-____', 'Brightline Technologies · Human Resources　No. LR-2026-____'), { color: C.sub, size: 10 }), { align: 'center' });
      await d.p('');
      await d.table([
        [L(t('申请人', 'Employee')), '', L(t('工号', 'Staff ID')), ''],
        [L(t('部门', 'Department')), '', L(t('岗位', 'Position')), ''],
        [L(t('请假类型', 'Leave type')), { html: check(t('年假', 'Annual'), t('事假', 'Personal'), t('病假', 'Sick'), t('婚假', 'Marriage'), t('产假 / 陪产假', 'Parental'), t('其他', 'Other')), colspan: 3 }, null, null],
        [L(t('开始时间', 'From')), esc(t('　　年　　月　　日　　午', '____ / ____ / ________  am / pm')), L(t('结束时间', 'To')), esc(t('　　年　　月　　日　　午', '____ / ____ / ________  am / pm'))],
        [L(t('请假天数', 'Days')), '', L(t('剩余年假', 'Annual leave left')), ''],
        [L(t('请假事由', 'Reason')), { html: '', colspan: 3, valign: 'top' }, null, null],
        [L(t('工作交接', 'Handover')), { html: esc(t('交接人：　　　　　　　交接内容：', 'Covered by:　　　　　　Details:')), colspan: 3, valign: 'top' }, null, null],
        [L(t('申请人签字', 'Employee signature')), esc(t('日期：', 'Date:')), L(t('紧急联系方式', 'Emergency contact')), ''],
        [L(t('部门负责人意见', 'Manager')), { html: esc(t('☐ 同意　☐ 不同意', '☐ Approved　☐ Declined')), valign: 'top' }, L(t('人力资源部意见', 'Human Resources')), { html: esc(t('☐ 同意　☐ 不同意', '☐ Approved　☐ Declined')), valign: 'top' }],
        [L(t('备注', 'Notes')), { html: span(t('1. 请假 1 天以内由部门负责人审批，3 天以上需分管副总审批。2. 病假需附医院证明。3. 本单一式两份，人力资源部与部门各存一份。', '1. Up to one day is approved by the manager; more than three days also needs the division head. 2. Sick leave needs a medical note. 3. Two copies: one for HR, one for the department.'), { size: 9.5, color: C.sub }), colspan: 3 }, null, null],
      ], { widths: ['3.2cm', '5cm', '3.2cm', '5cm'], borders: 'all', borderColor: C.line, heights: [1, 1, 1, 1, 1, 3.2, 2.2, 1, 3.4, 2] });
    },
  },
];
