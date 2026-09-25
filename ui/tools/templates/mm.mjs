// Mind map templates for the new-file gallery (see lib.mjs): tree([centre, …branches]), a branch being [text, …children].
export const cats = [['思考', 'Thinking'], ['计划', 'Planning'], ['学习', 'Study']];

export default [
  {
    id: 'project-plan', cat: '计划', name: ['项目规划', 'Project Plan'],
    build: (m, t) => m.tree([t('新版官网上线', 'New Website Launch'),
      [t('目标', 'Goals'), t('11 月 30 日正式上线', 'Live on November 30'), t('访问转化率提升 20%', 'Lift conversion by 20%'), t('移动端体验评分 4.5 以上', 'Mobile experience score of 4.5+')],
      [t('里程碑', 'Milestones'), t('10 月 10 日　设计定稿', 'Oct 10 · Design final'), t('10 月 31 日　开发完成', 'Oct 31 · Build complete'), t('11 月 20 日　测试验收', 'Nov 20 · Testing signed off'), t('11 月 30 日　上线', 'Nov 30 · Launch')],
      [t('分工', 'Owners'), t('产品：需求与验收', 'Product: scope and sign-off'), t('设计：视觉与交互', 'Design: visuals and flows'), t('开发：前端与后台', 'Engineering: front end and CMS'), t('市场：内容与推广', 'Marketing: content and promotion')],
      [t('风险', 'Risks'), t('内容迁移量大', 'A lot of content to migrate'), t('第三方接口不稳定', 'Unreliable third-party APIs'), t('测试时间偏紧', 'Tight testing window')],
      [t('资源', 'Resources'), t('预算 30 万元', 'Budget: $42,000'), t('团队 8 人', 'Team of eight'), t('外包：插画与视频', 'Outsourced: illustration and video')],
    ]),
  },
  {
    id: 'book-notes', cat: '学习', name: ['读书笔记', 'Book Notes'],
    build: (m, t) => m.tree([t('《深度工作》读书笔记', 'Deep Work · Reading Notes'),
      [{ text: t('作者与背景', 'Author and context'), side: 'right' }, t('卡尔·纽波特，计算机科学教授', 'Cal Newport, professor of computer science'), t('2016 年出版，讨论注意力经济下的工作方式', 'Published in 2016, on working in an attention economy')],
      [{ text: t('核心观点', 'Core ideas'), side: 'right' }, t('深度工作：无干扰地专注于认知要求高的任务', 'Deep work: focused effort on cognitively demanding tasks, without distraction'), t('浮浅工作：随时可做、容易被替代的事务', 'Shallow work: logistical tasks anyone could do, anytime'), t('深度工作的能力越来越稀缺，也越来越有价值', 'The ability to work deeply is getting rarer, and more valuable')],
      [{ text: t('四条原则', 'Four rules'), side: 'right' }, [t('工作要深入', 'Work deeply'), t('固定时段、固定地点、固定仪式', 'A set time, a set place, a set ritual')], [t('拥抱无聊', 'Embrace boredom'), t('不在每个空隙里掏手机', 'Do not fill every gap with the phone')], [t('远离社交媒体', 'Quit social media'), t('用「工匠思维」选择工具', 'Choose tools the way a craftsman does')], [t('摒弃浮浅', 'Drain the shallows'), t('提前规划每一天，给浮浅工作设上限', 'Plan every day in advance; cap shallow work')]],
      [{ text: t('金句摘录', 'Quotes'), side: 'left' }, t('「我会活在专注的世界里，因为专注让生活更好。」', '“I’ll live the focused life, because it’s the best kind there is.”'), t('「忙碌不等于生产力。」', '“Busyness is not productivity.”')],
      [{ text: t('我的思考', 'My thoughts'), side: 'left' }, t('我的深度时段其实在早上 7–9 点', 'My deep hours are really 7 to 9 in the morning'), t('会议之间的碎片时间只适合浮浅工作', 'The gaps between meetings only suit shallow work'), t('「关闭仪式」值得试：下班前写明天的三件事', 'Worth trying: a shutdown ritual that writes down tomorrow’s three things')],
      [{ text: t('行动清单', 'Actions'), side: 'left', icon: 'flag' }, t('每天早上封闭 90 分钟做最难的事', 'Block 90 minutes every morning for the hardest task'), t('手机通知只留电话与日历', 'Only calls and calendar may notify me'), t('周五回顾：本周深度工作小时数', 'Friday review: hours of deep work this week')],
    ]),
  },
  {
    id: 'swot', cat: '思考', name: ['SWOT 分析', 'SWOT Analysis'],
    build: (m, t) => m.tree([t('新产品上线 · SWOT', 'Product Launch · SWOT'),
      [{ text: t('优势 Strengths', 'Strengths'), side: 'right', fill: 'EEF2F7', color: '1F3A5F' }, t('技术领先：实时协同延迟行业最低', 'Technical lead: the lowest sync latency in the category'), t('现有 12 万活跃用户可直接触达', 'A direct line to 120,000 active users'), t('团队有两次成功发布的经验', 'A team that has shipped two launches before')],
      [{ text: t('劣势 Weaknesses', 'Weaknesses'), side: 'right', fill: 'F6EEE8', color: 'A4492F' }, t('品牌知名度低于头部竞品', 'Less brand recognition than the leaders'), t('销售团队小，渠道覆盖不足', 'A small sales team with thin channel coverage'), t('移动端体验落后桌面端', 'Mobile lags behind desktop')],
      [{ text: t('机会 Opportunities', 'Opportunities'), side: 'left', fill: 'F1EEF5', color: '4B3B63' }, t('中小团队数字化需求快速增长', 'Fast-growing demand from small teams'), t('竞品涨价，用户在寻找替代', 'Competitors raised prices; users are looking around'), t('AI 功能成为新的选购标准', 'AI features are becoming a buying criterion')],
      [{ text: t('威胁 Threats', 'Threats'), side: 'left', fill: 'F2F2F4', color: '1D1D1F' }, t('大厂可能捆绑免费同类功能', 'A large vendor could bundle a free equivalent'), t('数据合规要求趋严', 'Stricter data-compliance requirements'), t('获客成本持续上升', 'Rising acquisition costs')],
      [{ text: t('策略', 'Strategy'), icon: 'idea' }, t('SO：用现有用户做口碑传播，主打 AI 与速度', 'SO: word of mouth from current users, led by AI and speed'), t('WO：与渠道伙伴合作补足销售', 'WO: partner channels to make up for the small sales team'), t('ST：以数据安全认证建立信任', 'ST: build trust with security certifications'), t('WT：先聚焦桌面端优势市场', 'WT: focus first on the desktop segment we already win')],
    ]),
  },
  {
    id: 'okr', cat: '计划', name: ['季度 OKR', 'Quarterly OKRs'],
    build: (m, t) => m.tree([t('2026 Q4 OKR · 产品部', '2026 Q4 OKRs · Product'),
      [{ text: t('O1　让新用户第一周就用起来', 'O1 · New users get going in week one'), side: 'right', icon: 'flag' }, t('KR1　首周完成一次多人协作的用户占比从 22% 到 40%', 'KR1 · Users who collaborate with someone in week one: 22% → 40%'), t('KR2　次月留存率从 38% 到 45%', 'KR2 · Second-month retention: 38% → 45%'), t('KR3　新用户引导流程完成率 70% 以上', 'KR3 · Onboarding completion above 70%')],
      [{ text: t('O2　把稳定性做成口碑', 'O2 · Make reliability a selling point'), side: 'right' }, t('KR1　服务可用性 99.9%', 'KR1 · 99.9% availability'), t('KR2　100 页文档打开时间 P95 < 1 秒', 'KR2 · P95 open time for a 100-page document under one second'), t('KR3　故障平均恢复时间 < 15 分钟', 'KR3 · Mean time to recovery under 15 minutes')],
      [{ text: t('O3　让团队版值得付费', 'O3 · Make the team plan worth paying for'), side: 'left' }, t('KR1　团队版试用转付费 12%', 'KR1 · 12% of team trials convert'), t('KR2　上线权限与审计日志', 'KR2 · Ship permissions and the audit log'), t('KR3　付费团队 NPS ≥ 45', 'KR3 · NPS of paying teams at 45 or above')],
      [{ text: t('复盘节奏', 'Check-ins'), side: 'left' }, t('每周一 30 分钟进度同步', 'Thirty minutes every Monday'), t('每月底打分：0.0–1.0', 'Score each KR 0.0–1.0 at month end'), t('12 月 20 日季度复盘', 'Quarter review on December 20')],
    ]),
  },
];
