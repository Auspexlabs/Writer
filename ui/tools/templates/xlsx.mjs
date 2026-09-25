// Excel templates for the new-file gallery (see lib.mjs). Every total, balance, share, rank and status is a formula.
import { NAVY, CLAY, INK, SLATE, PLUM, OCHRE, colName } from './lib.mjs';

export const cats = [['个人生活', 'Personal'], ['财务', 'Financial'], ['人事行政', 'HR & Admin'], ['项目管理', 'Projects'], ['销售运营', 'Sales & Ops'], ['学习', 'Study']];

const WHITE = 'FFFFFF', MONEY = '#,##0.00', INT = '#,##0', PCT = '0.0%';
/** A sheet title in row 1 across the given columns. */
const title = async (b, C, text, cols, size = 16) => {
  await b.cell('A1', { value: text });
  await b.style('A1', { bold: true, size: size + 'pt', color: C.acc });
  await b.merge(`A1:${cols}1`);
  await b.layout({ heights: { 1: 30 } });
};
/** A header row: accent fill, white bold text, centred. */
const header = (b, C, ref) => b.style(ref, { bold: true, fill: C.acc, color: WHITE, align: 'center', valign: 'middle' });
/** Thin light lines on every side of every cell. */
const grid = (b, C, ref) => b.style(ref, { border: 'thin', borderColor: C.line });
/** A strip of headline numbers: labels in row r (small, grey), formulas in row r + 1 (large, accent), each two columns wide. */
const colIdx = s => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
async function kpis(b, C, r, items, col = 'A') {
  let ci = colIdx(col);
  for (const [label, formula, format] of items) {
    const c = colName(ci), next = colName(ci + 1);
    await b.cell(`${c}${r}`, { value: label });
    await b.style(`${c}${r}`, { color: C.sub, size: '10pt' });
    await b.fx(`${c}${r + 1}`, formula.replace(/^=/, ''), { bold: true, size: '16pt', color: C.acc, format, align: 'left' });
    await b.merge(`${c}${r}:${next}${r}`, `${c}${r + 1}:${next}${r + 1}`);
    ci += 2;
  }
  await b.layout({ heights: { [r + 1]: 26 } });
}
const DATE = t => t('m"月"d"日"', 'mmm d');

export default [
  {
    id: 'personal-ledger', cat: '个人生活', name: ['个人记账', 'Personal Ledger'],
    async build(b, t) {
      const C = NAVY, first = 7, last = 40;
      await b.sheet(t('9月账本', 'September'));
      await title(b, C, t('2026 年 9 月 · 个人记账', 'Personal Ledger · September 2026'), 'G');
      await kpis(b, C, 3, [[t('本月收入', 'Income'), `=SUM(E${first}:E${last})`, MONEY], [t('本月支出', 'Spending'), `=SUM(F${first}:F${last})`, MONEY], [t('本月结余', 'Balance'), '=A4-C4', MONEY]]);
      await b.cell('G3', { value: t('支出占比', 'Spent') }); await b.style('G3', { color: C.sub, size: '10pt' });
      await b.fx('G4', 'IFERROR(C4/A4,0)', { bold: true, size: '16pt', color: C.acc, format: PCT, align: 'left' });
      await b.values('A6:G6', [[t('日期', 'Date'), t('类别', 'Category'), t('说明', 'Description'), t('账户', 'Account'), t('收入', 'Income'), t('支出', 'Spending'), t('余额', 'Balance')]]);
      await header(b, C, 'A6:G6');
      const W = t('微信', 'Card'), A = t('支付宝', 'Cash'), K = t('银行卡', 'Bank');
      const rows = [
        ['2026-09-01', t('收入', 'Income'), t('9 月工资', 'September salary'), K, 12800, null],
        ['2026-09-01', t('居住', 'Home'), t('房租', 'Rent'), K, null, 3200],
        ['2026-09-02', t('餐饮', 'Food'), t('超市采购', 'Groceries'), A, null, 286.5],
        ['2026-09-03', t('交通', 'Transport'), t('地铁卡充值', 'Transit card top-up'), W, null, 100],
        ['2026-09-05', t('餐饮', 'Food'), t('朋友聚餐', 'Dinner with friends'), W, null, 238],
        ['2026-09-07', t('购物', 'Shopping'), t('秋季外套', 'Autumn jacket'), A, null, 459],
        ['2026-09-10', t('居住', 'Home'), t('水电燃气', 'Utilities'), K, null, 312.4],
        ['2026-09-12', t('娱乐', 'Leisure'), t('电影与展览', 'Cinema and a show'), W, null, 168],
        ['2026-09-15', t('收入', 'Income'), t('兼职稿费', 'Freelance article'), K, 1500, null],
        ['2026-09-16', t('医疗', 'Health'), t('体检', 'Health check'), A, null, 420],
        ['2026-09-18', t('餐饮', 'Food'), t('工作日午餐', 'Weekday lunches'), W, null, 356],
        ['2026-09-21', t('学习', 'Learning'), t('线上课程', 'Online course'), A, null, 299],
        ['2026-09-24', t('交通', 'Transport'), t('打车', 'Taxi'), W, null, 64.8],
        ['2026-09-27', t('购物', 'Shopping'), t('日用品', 'Household items'), A, null, 132.9],
      ];
      await b.values(`A${first}:F${first + rows.length - 1}`, rows);
      for (let r = first; r <= last; r++) await b.fx(`G${r}`, `IF(AND(E${r}="",F${r}=""),"",SUM(E$${first}:E${r})-SUM(F$${first}:F${r}))`);
      await b.style(`A${first}:A${last}`, { format: DATE(t), align: 'left' });
      await b.style(`E${first}:G${last}`, { format: MONEY });
      await grid(b, C, `A${first}:G${last}`);
      await b.style(`E${first}:E${last}`, { color: SLATE.acc });
      await b.style(`F${first}:F${last}`, { color: CLAY.acc });
      // spending by category, with a doughnut of it
      const cats = [t('居住', 'Home'), t('餐饮', 'Food'), t('交通', 'Transport'), t('购物', 'Shopping'), t('娱乐', 'Leisure'), t('医疗', 'Health'), t('学习', 'Learning')];
      await b.values('I6:K6', [[t('类别', 'Category'), t('支出', 'Spending'), t('占比', 'Share')]]);
      await header(b, C, 'I6:K6');
      await b.values(`I7:K${6 + cats.length}`, cats.map((c, i) => [c, `=SUMIF($B$${first}:$B$${last},I${7 + i},$F$${first}:$F$${last})`, `=IFERROR(J${7 + i}/SUM($J$7:$J$${6 + cats.length}),0)`]));
      const tot = 7 + cats.length;
      await b.values(`I${tot}:K${tot}`, [[t('合计', 'Total'), `=SUM(J7:J${tot - 1})`, `=SUM(K7:K${tot - 1})`]]);
      await b.style(`J7:J${tot}`, { format: MONEY });
      await b.style(`K7:K${tot}`, { format: PCT });
      await grid(b, C, `I7:K${tot}`);
      await b.style(`I${tot}:K${tot}`, { bold: true, fill: C.soft });
      await b.chart({ type: 'doughnut', title: t('支出构成', 'Where it went'), categories: `I7:I${tot - 1}`, series: [{ name: 'J6', values: `J7:J${tot - 1}` }], legend: 'right', x: 16.4, y: 9.4, w: 9, h: 6.5 });
      await b.layout({ widths: { A: 10, B: 11, C: 18, D: 9, E: 11, F: 11, G: 12, H: 3, I: 10, J: 11, K: 8 }, heights: { 6: 22 }, freeze: 'A7' });
    },
  },
  {
    id: 'budget', cat: '财务', name: ['月度预算', 'Monthly Budget'],
    async build(b, t) {
      const C = SLATE;
      await b.sheet(t('10月预算', 'October'));
      await title(b, C, t('2026 年 10 月 · 家庭预算', 'Household Budget · October 2026'), 'G');
      await kpis(b, C, 3, [[t('实际收入', 'Income'), '=D10', MONEY], [t('实际支出', 'Spending'), '=D21', MONEY], [t('结余', 'Left over'), '=A4-C4', MONEY]]);
      await b.cell('G3', { value: t('储蓄率', 'Saved') }); await b.style('G3', { color: C.sub, size: '10pt' });
      await b.fx('G4', 'IFERROR((A4-C4+D20)/A4,0)', { bold: true, size: '16pt', color: C.acc, format: PCT, align: 'left' });
      await b.values('A6:G6', [[t('类别', 'Group'), t('项目', 'Item'), t('预算', 'Budget'), t('实际', 'Actual'), t('差额', 'Difference'), t('执行率', 'Used'), t('备注', 'Notes')]]);
      await header(b, C, 'A6:G6');
      const inc = t('收入', 'Income'), exp = t('支出', 'Spending');
      const rows = [
        [inc, t('工资', 'Salary'), 15000, 15000, ''],
        [inc, t('奖金', 'Bonus'), 2000, 1500, t('季度奖延后一部分', 'Part of the quarterly bonus is late')],
        [inc, t('其他收入', 'Other'), 500, 800, ''],
        [{ v: t('收入小计', 'Income total') }],
        [exp, t('房租 / 房贷', 'Rent / mortgage'), 5000, 5000, ''],
        [exp, t('餐饮', 'Food'), 3000, 3260, ''],
        [exp, t('交通', 'Transport'), 800, 720, ''],
        [exp, t('通讯', 'Phone & internet'), 200, 200, ''],
        [exp, t('水电燃气', 'Utilities'), 400, 380, ''],
        [exp, t('购物', 'Shopping'), 1500, 1880, t('换季衣物', 'Autumn clothes')],
        [exp, t('娱乐', 'Leisure'), 600, 450, ''],
        [exp, t('教育', 'Learning'), 800, 800, ''],
        [exp, t('医疗', 'Health'), 300, 120, ''],
        [exp, t('储蓄 / 投资', 'Savings'), 3000, 3000, t('计入支出，结余之外', 'Counted as spending, on top of what is left')],
        [{ v: t('支出小计', 'Spending total') }],
      ];
      let r = 7;
      for (const row of rows) {
        if (row[0].v) { // a subtotal of the rows above it
          const from = row[0].v === rows[3][0].v ? 7 : 11;
          await b.values(`A${r}:F${r}`, [[row[0].v, '', `=SUM(C${from}:C${r - 1})`, `=SUM(D${from}:D${r - 1})`, `=C${r}-D${r}`, `=IFERROR(D${r}/C${r},0)`]]);
          await b.merge(`A${r}:B${r}`);
          await b.style(`A${r}:G${r}`, { bold: true, fill: C.soft });
        } else await b.values(`A${r}:G${r}`, [[row[0], row[1], row[2], row[3], `=C${r}-D${r}`, `=IFERROR(D${r}/C${r},0)`, row[4]]]);
        r++;
      }
      await b.style('C7:E21', { format: MONEY });
      await b.style('F7:F21', { format: PCT });
      await b.style('A7:A21', { color: C.sub });
      await grid(b, C, 'A7:G21');
      await b.chart({ type: 'column', title: t('支出：预算与实际', 'Spending: budget and actual'), categories: 'B11:B20', series: [{ name: 'C6', values: 'C11:C20' }, { name: 'D6', values: 'D11:D20' }], legend: 'bottom', x: 24.6, y: 3.2, w: 13, h: 8.5 });
      await b.layout({ widths: { A: 9, B: 16, C: 11, D: 11, E: 11, F: 9, G: 24 }, heights: { 6: 22 }, freeze: 'A7' });
    },
  },
  {
    id: 'attendance', cat: '人事行政', name: ['考勤表', 'Attendance Sheet'],
    async build(b, t) {
      const C = NAVY, DAYS = 30, first = 5, names = [t('李明', 'Alex Chen'), t('王芳', 'Jordan Lee'), t('张伟', 'Sam Taylor'), t('陈静', 'Morgan Reed'), t('刘洋', 'Casey Kim'), t('赵磊', 'Riley Park'), t('孙丽', 'Taylor Brooks'), t('周杰', 'Jamie Cruz'), t('吴敏', 'Dana Wells'), t('郑浩', 'Avery Stone')];
      const depts = [t('市场部', 'Marketing'), t('产品部', 'Product'), t('研发部', 'Engineering'), t('行政部', 'Admin'), t('财务部', 'Finance')];
      const M = { on: t('✓', 'P'), off: t('休', 'W'), leave: t('假', 'L'), late: t('迟', 'LT'), absent: t('缺', 'A') };
      const day0 = colName(3), dayN = colName(2 + DAYS), sum0 = colName(3 + DAYS), sumN = colName(6 + DAYS), last = first + names.length - 1;
      const wd = d => (d + 1) % 7; // 2026-09-01 is a Tuesday: 0 = Sunday
      const wkNames = t(['日', '一', '二', '三', '四', '五', '六'], ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa']);
      await b.sheet(t('9月考勤', 'September'));
      await title(b, C, t('2026 年 9 月 · 员工考勤表', 'Attendance · September 2026'), sumN);
      await b.values('A3:C3', [[t('工号', 'ID'), t('姓名', 'Name'), t('部门', 'Team')]]);
      await b.values(`${day0}3:${dayN}4`, [Array.from({ length: DAYS }, (_, i) => i + 1), Array.from({ length: DAYS }, (_, i) => wkNames[wd(i + 1)])]);
      await b.values(`${sum0}3:${sumN}3`, [[t('出勤', 'Present'), t('请假', 'Leave'), t('迟到', 'Late'), t('缺勤', 'Absent')]]);
      await b.merge('A3:A4', 'B3:B4', 'C3:C4', ...[0, 1, 2, 3].map(i => `${colName(3 + DAYS + i)}3:${colName(3 + DAYS + i)}4`));
      await header(b, C, `A3:${sumN}4`);
      const marks = (i, d) => { if (wd(d) === 0 || wd(d) === 6) return M.off; const k = (i * 7 + d * 3) % 29; return k === 4 ? M.leave : k === 11 ? M.late : k === 20 && i % 3 === 0 ? M.absent : M.on; };
      await b.values(`A${first}:${sumN}${last}`, names.map((n, i) => {
        const r = first + i;
        return [String(1001 + i), n, depts[i % depts.length], ...Array.from({ length: DAYS }, (_, d) => marks(i, d + 1)),
          `=COUNTIF(${day0}${r}:${dayN}${r},"${M.on}")+COUNTIF(${day0}${r}:${dayN}${r},"${M.late}")`, `=COUNTIF(${day0}${r}:${dayN}${r},"${M.leave}")`, `=COUNTIF(${day0}${r}:${dayN}${r},"${M.late}")`, `=COUNTIF(${day0}${r}:${dayN}${r},"${M.absent}")`];
      }));
      await b.values(`A${last + 1}:C${last + 1}`, [[t('合计', 'Total'), '', '']]);
      await b.merge(`A${last + 1}:C${last + 1}`);
      await b.values(`${sum0}${last + 1}:${sumN}${last + 1}`, [[0, 1, 2, 3].map(i => `=SUM(${colName(3 + DAYS + i)}${first}:${colName(3 + DAYS + i)}${last})`)]);
      await b.style(`A${last + 1}:${sumN}${last + 1}`, { bold: true, fill: C.soft });
      await b.style(`${day0}${first}:${sumN}${last + 1}`, { align: 'center' });
      await grid(b, C, `A${first}:${sumN}${last + 1}`);
      for (let d = 1; d <= DAYS; d++) if (wd(d) === 0 || wd(d) === 6) await b.style(`${colName(2 + d)}${first}:${colName(2 + d)}${last}`, { fill: 'F2F2F4', color: C.sub });
      await b.cell(`A${last + 3}`, { value: t(`说明：✓ 出勤　休 周末　假 请假　迟 迟到　缺 缺勤；出勤天数含迟到。`, 'P present · W weekend · L leave · LT late · A absent. Days present include late days.') });
      await b.style(`A${last + 3}`, { color: C.sub, size: '10pt' });
      const widths = { A: 7, B: 10, C: 10 };
      for (let d = 1; d <= DAYS; d++) widths[colName(2 + d)] = 3.6;
      for (let i = 0; i < 4; i++) widths[colName(3 + DAYS + i)] = 7;
      await b.layout({ widths, heights: { 3: 20, 4: 20 }, freeze: `${day0}${first}` });
    },
  },
  {
    id: 'inventory', cat: '销售运营', name: ['库存表', 'Inventory'],
    async build(b, t) {
      const C = INK, first = 7;
      await b.sheet(t('库存', 'Stock'));
      await title(b, C, t('办公用品库存表 · 2026 年 9 月', 'Office Supplies Inventory · September 2026'), 'L');
      await kpis(b, C, 3, [[t('库存金额', 'Stock value'), `=SUM(J${first}:J40)`, MONEY], [t('品目数', 'Items'), `=COUNTA(B${first}:B40)`, INT], [t('需补货', 'To reorder'), `=COUNTIF(L${first}:L40,"${t('需补货', 'Reorder')}")`, INT]]);
      await b.values('A6:L6', [[t('编号', 'Code'), t('名称', 'Item'), t('类别', 'Category'), t('单位', 'Unit'), t('期初', 'Opening'), t('入库', 'In'), t('出库', 'Out'), t('期末', 'Closing'), t('单价', 'Unit price'), t('库存金额', 'Value'), t('安全库存', 'Min. stock'), t('状态', 'Status')]]);
      await header(b, C, 'A6:L6');
      const P = t('纸品', 'Paper'), S = t('文具', 'Stationery'), E = t('设备', 'Equipment'), H = t('耗材', 'Consumables');
      const items = [
        ['A001', t('A4 复印纸', 'A4 copy paper'), P, t('箱', 'box'), 24, 20, 31, 45, 10],
        ['A002', t('便利贴', 'Sticky notes'), P, t('本', 'pad'), 60, 0, 38, 4.5, 20],
        ['A003', t('档案盒', 'Archive boxes'), P, t('个', 'each'), 35, 50, 42, 6.8, 20],
        ['B001', t('签字笔（黑）', 'Gel pens, black'), S, t('支', 'each'), 120, 100, 168, 2.5, 60],
        ['B002', t('白板笔', 'Whiteboard markers'), S, t('支', 'each'), 40, 0, 28, 3.2, 20],
        ['B003', t('订书机', 'Staplers'), S, t('个', 'each'), 12, 0, 3, 28, 5],
        ['B004', t('文件夹', 'Folders'), S, t('个', 'each'), 80, 40, 66, 5.5, 30],
        ['B005', t('透明胶带', 'Tape'), S, t('卷', 'roll'), 30, 24, 39, 4, 15],
        ['C001', t('硒鼓', 'Toner cartridges'), H, t('个', 'each'), 6, 4, 7, 320, 3],
        ['C002', t('5 号电池', 'AA batteries'), H, t('节', 'each'), 48, 0, 36, 1.8, 24],
        ['C003', t('打印机墨盒', 'Ink cartridges'), H, t('个', 'each'), 8, 6, 9, 145, 4],
        ['D001', t('键盘', 'Keyboards'), E, t('个', 'each'), 10, 5, 4, 119, 4],
        ['D002', t('鼠标', 'Mice'), E, t('个', 'each'), 14, 0, 6, 69, 4],
        ['D003', t('显示器', 'Monitors'), E, t('台', 'each'), 4, 6, 7, 1280, 2],
        ['D004', t('网线（3 米）', 'Network cables, 3 m'), E, t('根', 'each'), 25, 0, 12, 12, 10],
      ];
      const last = first + items.length - 1;
      await b.values(`A${first}:L${last}`, items.map((x, i) => { const r = first + i; return [...x.slice(0, 7), `=E${r}+F${r}-G${r}`, x[7], `=H${r}*I${r}`, x[8], `=IF(H${r}<=K${r},"${t('需补货', 'Reorder')}","${t('正常', 'OK')}")`]; }));
      await b.values(`A${last + 1}:L${last + 1}`, [[t('合计', 'Total'), '', '', '', `=SUM(E${first}:E${last})`, `=SUM(F${first}:F${last})`, `=SUM(G${first}:G${last})`, `=SUM(H${first}:H${last})`, '', `=SUM(J${first}:J${last})`, '', '']]);
      await b.merge(`A${last + 1}:D${last + 1}`);
      await b.style(`A${last + 1}:L${last + 1}`, { bold: true, fill: C.soft });
      await b.style(`E${first}:H${last + 1}`, { format: INT, align: 'right' });
      await b.style(`K${first}:K${last}`, { format: INT, align: 'right' });
      await b.style(`I${first}:J${last + 1}`, { format: MONEY });
      await b.style(`D${first}:D${last}`, { align: 'center' });
      await b.style(`L${first}:L${last}`, { align: 'center', color: CLAY.acc });
      await grid(b, C, `A${first}:L${last + 1}`);
      await b.layout({ widths: { A: 8, B: 18, C: 9, D: 6, E: 8, F: 8, G: 8, H: 8, I: 10, J: 12, K: 9, L: 9 }, heights: { 6: 22 }, freeze: `C${first}`, filter: `A6:L${last}` });
    },
  },
  {
    id: 'project-tracker', cat: '项目管理', name: ['项目进度表', 'Project Tracker'],
    async build(b, t) {
      const C = PLUM, first = 7, WEEKS = 12, w0 = 10; // week columns from K
      const done = t('已完成', 'Done'), todo = t('未开始', 'Not started'), late = t('逾期', 'Late'), going = t('进行中', 'In progress');
      await b.sheet(t('进度', 'Schedule'));
      await title(b, C, t('新版官网改版 · 项目进度表', 'Website Redesign · Project Tracker'), colName(w0 + WEEKS - 1));
      await b.values('A3:B3', [[t('统计日期', 'As of'), '2026-10-15']]);
      await b.style('A3', { color: C.sub, size: '10pt' }); await b.style('B3', { bold: true, format: 'yyyy-mm-dd', align: 'left' });
      await kpis(b, C, 3, [[t('任务数', 'Tasks'), `=COUNTA(B${first}:B40)`, INT], [t('已完成', 'Done'), `=COUNTIF(I${first}:I40,"${done}")`, INT], [t('逾期', 'Late'), `=COUNTIF(I${first}:I40,"${late}")`, INT], [t('整体进度', 'Overall'), `=AVERAGE(H${first}:H40)`, PCT]], 'D');
      await b.values('A6:J6', [[t('编号', '#'), t('任务', 'Task'), t('阶段', 'Phase'), t('负责人', 'Owner'), t('开始', 'Start'), t('结束', 'End'), t('工期', 'Days'), t('进度', 'Progress'), t('状态', 'Status'), '']]);
      await b.values(`${colName(w0)}6:${colName(w0 + WEEKS - 1)}6`, [Array.from({ length: WEEKS }, (_, i) => { const d = new Date(Date.UTC(2026, 8, 7 + i * 7)); return d.toISOString().slice(0, 10); })]);
      await header(b, C, `A6:${colName(w0 + WEEKS - 1)}6`);
      await b.style(`${colName(w0)}6:${colName(w0 + WEEKS - 1)}6`, { format: t('m/d', 'm/d'), size: '9pt' });
      const D = t('需求', 'Discovery'), G = t('设计', 'Design'), B = t('开发', 'Build'), T = t('测试上线', 'Launch');
      const tasks = [
        [t('需求调研', 'User research'), D, t('林晓', 'Alex Chen'), '2026-09-28', '2026-10-10', 1],
        [t('信息架构', 'Information architecture'), D, t('林晓', 'Alex Chen'), '2026-10-05', '2026-10-12', 1],
        [t('视觉方案', 'Visual direction'), G, t('王芳', 'Jordan Lee'), '2026-10-11', '2026-10-20', 0.8],
        [t('页面设计', 'Page design'), G, t('王芳', 'Jordan Lee'), '2026-10-13', '2026-10-31', 0.45],
        [t('设计评审', 'Design review'), G, t('张伟', 'Sam Taylor'), '2026-10-08', '2026-10-14', 0.6],
        [t('前端开发', 'Front-end build'), B, t('刘洋', 'Casey Kim'), '2026-10-20', '2026-11-12', 0.2],
        [t('CMS 接入', 'CMS integration'), B, t('刘洋', 'Casey Kim'), '2026-10-27', '2026-11-08', 0.1],
        [t('内容迁移', 'Content migration'), B, t('陈静', 'Morgan Reed'), '2026-11-01', '2026-11-15', 0],
        [t('功能测试', 'QA'), T, t('陈静', 'Morgan Reed'), '2026-11-10', '2026-11-22', 0],
        [t('性能优化', 'Performance tuning'), T, t('刘洋', 'Casey Kim'), '2026-11-16', '2026-11-25', 0],
        [t('灰度发布', 'Staged rollout'), T, t('林晓', 'Alex Chen'), '2026-11-23', '2026-11-28', 0],
        [t('正式上线', 'Launch'), T, t('林晓', 'Alex Chen'), '2026-11-30', '2026-11-30', 0],
      ];
      const last = first + tasks.length - 1;
      await b.values(`A${first}:${colName(w0 + WEEKS - 1)}${last}`, tasks.map((x, i) => {
        const r = first + i;
        return [i + 1, x[0], x[1], x[2], x[3], x[4], `=F${r}-E${r}+1`, x[5], `=IF(H${r}>=1,"${done}",IF(E${r}>$B$3,"${todo}",IF(F${r}<$B$3,"${late}","${going}")))`, '',
          ...Array.from({ length: WEEKS }, (_, k) => { const c = colName(w0 + k); return `=IF(AND(${c}$6<=$F${r},${c}$6+6>=$E${r}),"■","")`; })];
      }));
      await b.style(`E${first}:F${last}`, { format: t('m月d日', 'mmm d'), align: 'center' });
      await b.style(`A${first}:A${last}`, { align: 'center', color: C.sub });
      await b.style(`G${first}:G${last}`, { align: 'center' });
      await b.style(`H${first}:H${last}`, { format: '0%', align: 'center' });
      await b.style(`I${first}:I${last}`, { align: 'center' });
      await b.style(`${colName(w0)}${first}:${colName(w0 + WEEKS - 1)}${last}`, { align: 'center', color: C.acc });
      await grid(b, C, `A${first}:I${last}`);
      await b.style(`${colName(w0)}${first}:${colName(w0 + WEEKS - 1)}${last}`, { border: 'thin', borderColor: 'EDEAF2' });
      const widths = { A: 5, B: 18, C: 9, D: 10, E: 9, F: 9, G: 6, H: 7, I: 9, J: 2 };
      for (let k = 0; k < WEEKS; k++) widths[colName(w0 + k)] = 5.2;
      await b.layout({ widths, heights: { 6: 22 }, freeze: `C${first}` });
    },
  },
  {
    id: 'sales-report', cat: '销售运营', name: ['销售报表', 'Sales Report'],
    async build(b, t) {
      const C = CLAY, first = 7;
      await b.sheet(t('2026销售', 'Sales 2026'));
      await title(b, C, t('2026 年销售报表（万元）', 'Sales Report 2026 ($ thousands)'), 'O');
      const months = t(['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'], ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']);
      const regions = [
        [t('华东', 'East'), [186, 142, 205, 198, 221, 240, 232, 228, 251, 268, 302, 335]],
        [t('华北', 'North'), [132, 98, 141, 150, 163, 171, 168, 175, 180, 194, 226, 248]],
        [t('华南', 'South'), [121, 90, 128, 139, 146, 158, 165, 160, 172, 181, 204, 230]],
        [t('西南', 'West'), [74, 55, 82, 88, 95, 101, 99, 108, 112, 120, 133, 150]],
        [t('华中', 'Central'), [66, 48, 70, 77, 84, 90, 88, 95, 101, 106, 122, 138]],
      ];
      const last = first + regions.length - 1, tot = last + 1;
      await kpis(b, C, 3, [[t('年度合计', 'Year total'), `=N${tot}`, INT], [t('月均', 'Monthly average'), `=AVERAGE(B${tot}:M${tot})`, INT], [t('最高月份', 'Best month'), `=INDEX(B6:M6,MATCH(MAX(B${tot}:M${tot}),B${tot}:M${tot},0))`, ''], [t('最佳区域', 'Best region'), `=INDEX(A${first}:A${last},MATCH(MAX(N${first}:N${last}),N${first}:N${last},0))`, '']]);
      await b.values('A6:O6', [[t('区域', 'Region'), ...months, t('合计', 'Total'), t('占比', 'Share')]]);
      await header(b, C, 'A6:O6');
      await b.values(`A${first}:O${last}`, regions.map(([name, v], i) => { const r = first + i; return [name, ...v, `=SUM(B${r}:M${r})`, `=IFERROR(N${r}/$N$${tot},0)`]; }));
      await b.values(`A${tot}:O${tot}`, [[t('合计', 'Total'), ...months.map((_, i) => `=SUM(${colName(1 + i)}${first}:${colName(1 + i)}${last})`), `=SUM(N${first}:N${last})`, `=SUM(O${first}:O${last})`]]);
      await b.values(`A${tot + 1}:M${tot + 1}`, [[t('环比', 'vs previous month'), '', ...months.slice(1).map((_, i) => `=IFERROR(${colName(2 + i)}${tot}/${colName(1 + i)}${tot}-1,"")`)]]);
      await b.style(`B${first}:N${tot}`, { format: INT });
      await b.style(`O${first}:O${tot}`, { format: PCT });
      await b.style(`A${tot}:O${tot}`, { bold: true, fill: C.soft });
      await b.style(`A${tot + 1}:M${tot + 1}`, { format: PCT, color: C.sub, size: '10pt' });
      await grid(b, C, `A${first}:O${tot}`);
      // who sold what, against target
      const p0 = tot + 4;
      await b.cell(`A${p0 - 1}`, { value: t('销售人员完成情况', 'By salesperson') }); await b.style(`A${p0 - 1}`, { bold: true, size: '12pt', color: C.acc });
      await b.values(`A${p0}:F${p0}`, [[t('姓名', 'Name'), t('区域', 'Region'), t('年度目标', 'Target'), t('实际', 'Actual'), t('完成率', 'Attainment'), t('排名', 'Rank')]]);
      await header(b, C, `A${p0}:F${p0}`);
      const people = [[t('李明', 'Alex Chen'), regions[0][0], 2600, `=N${first}`], [t('王芳', 'Jordan Lee'), regions[1][0], 1900, `=N${first + 1}`], [t('张伟', 'Sam Taylor'), regions[2][0], 1800, `=N${first + 2}`], [t('陈静', 'Morgan Reed'), regions[3][0], 1100, `=N${first + 3}`], [t('刘洋', 'Casey Kim'), regions[4][0], 1000, `=N${first + 4}`]];
      const pl = p0 + people.length;
      await b.values(`A${p0 + 1}:F${pl}`, people.map((x, i) => { const r = p0 + 1 + i; return [...x, `=IFERROR(D${r}/C${r},0)`, `=RANK(E${r},$E$${p0 + 1}:$E$${pl})`]; }));
      await b.style(`C${p0 + 1}:D${pl}`, { format: INT });
      await b.style(`E${p0 + 1}:E${pl}`, { format: PCT });
      await b.style(`F${p0 + 1}:F${pl}`, { align: 'center' });
      await grid(b, C, `A${p0}:F${pl}`);
      await b.chart({ type: 'line', title: t('月度销售趋势', 'Monthly sales'), categories: 'B6:M6', series: [{ name: `A${tot}`, values: `B${tot}:M${tot}` }], legend: 'none', x: 15.2, y: 8.6, w: 15, h: 7.5 });
      await b.chart({ type: 'bar', title: t('区域全年销售', 'Sales by region'), categories: `A${first}:A${last}`, series: [{ name: 'N6', values: `N${first}:N${last}` }], legend: 'none', x: 15.2, y: 16.6, w: 15, h: 6.5 });
      const widths = { A: 12 }; for (let i = 0; i < 12; i++) widths[colName(1 + i)] = 7; widths.N = 9; widths.O = 8;
      await b.layout({ widths, heights: { 6: 22 }, freeze: `B${first}` });
    },
  },
  {
    id: 'timetable', cat: '学习', name: ['课程表', 'Class Timetable'],
    async build(b, t) {
      const C = SLATE, first = 5;
      await b.sheet(t('课程表', 'Timetable'));
      await title(b, C, t('2026 年秋季学期 · 课程表', 'Class Timetable · Autumn 2026'), 'G');
      await b.cell('A2', { value: t('高二（3）班　　班主任：王老师', 'Year 11, Class 3 · Homeroom teacher: Ms. Lee') }); await b.style('A2', { color: C.sub, size: '10pt' }); await b.merge('A2:G2');
      await b.values('A4:G4', [[t('节次', 'Period'), t('时间', 'Time'), t('周一', 'Monday'), t('周二', 'Tuesday'), t('周三', 'Wednesday'), t('周四', 'Thursday'), t('周五', 'Friday')]]);
      await header(b, C, 'A4:G4');
      const S = { ch: t('语文', 'Literature'), ma: t('数学', 'Maths'), en: t('英语', 'English'), ph: t('物理', 'Physics'), ce: t('化学', 'Chemistry'), bi: t('生物', 'Biology'), hi: t('历史', 'History'), ge: t('地理', 'Geography'), pe: t('体育', 'PE'), mu: t('音乐', 'Music'), ar: t('美术', 'Art'), it: t('信息技术', 'Computing'), st: t('自习', 'Study'), cm: t('班会', 'Homeroom') };
      const rows = [
        ['1', '08:00–08:45', S.ch, S.ma, S.en, S.ma, S.ch],
        ['2', '08:55–09:40', S.ma, S.ph, S.ch, S.en, S.ma],
        ['3', '10:00–10:45', S.en, S.ch, S.ph, S.ce, S.en],
        ['4', '10:55–11:40', S.ph, S.en, S.ma, S.ph, S.bi],
        [{ v: t('午休　12:00–13:30', 'Lunch  12:00–13:30') }],
        ['5', '13:30–14:15', S.ce, S.hi, S.ge, S.ch, S.ph],
        ['6', '14:25–15:10', S.hi, S.bi, S.ce, S.it, S.ge],
        ['7', '15:30–16:15', S.pe, S.ar, S.st, S.pe, S.mu],
        ['8', '16:25–17:10', S.st, S.st, S.cm, S.st, S.st],
      ];
      let r = first;
      for (const row of rows) {
        if (row[0].v) { await b.cell(`A${r}`, { value: row[0].v }); await b.merge(`A${r}:G${r}`); await b.style(`A${r}:G${r}`, { fill: C.soft, color: C.sub, align: 'center', valign: 'middle' }); }
        else await b.values(`A${r}:G${r}`, [row]);
        r++;
      }
      const last = r - 1;
      await b.style(`A${first}:G${last}`, { align: 'center', valign: 'middle' });
      await b.style(`A${first}:B${last}`, { color: C.sub });
      await b.style(`A${first}:A${last}`, { bold: true, color: C.acc });
      await grid(b, C, `A${first}:G${last}`);
      const heights = { 4: 24 }; for (let i = first; i <= last; i++) heights[i] = 30;
      // lessons per subject: a count over the grid
      const subjects = [S.ch, S.ma, S.en, S.ph, S.ce, S.bi, S.hi, S.ge, S.pe, S.mu, S.ar, S.it, S.st, S.cm];
      await b.values('I4:J4', [[t('科目', 'Subject'), t('每周节数', 'Lessons / week')]]);
      await header(b, C, 'I4:J4');
      await b.values(`I5:J${4 + subjects.length}`, subjects.map((s, i) => [s, `=COUNTIF($C$${first}:$G$${last},I${5 + i})`]));
      await b.values(`I${5 + subjects.length}:J${5 + subjects.length}`, [[t('合计', 'Total'), `=SUM(J5:J${4 + subjects.length})`]]);
      await b.style(`I${5 + subjects.length}:J${5 + subjects.length}`, { bold: true, fill: C.soft });
      await b.style(`J5:J${5 + subjects.length}`, { align: 'center' });
      await grid(b, C, `I5:J${5 + subjects.length}`);
      await b.layout({ widths: { A: 6, B: 13, C: 12, D: 12, E: 12, F: 12, G: 12, H: 3, I: 11, J: 11 }, heights });
    },
  },
  {
    id: 'quotation', cat: '销售运营', name: ['报价单', 'Quotation'],
    async build(b, t) {
      const C = NAVY, CUR = t('¥#,##0.00', '$#,##0.00'), first = 10;
      await b.sheet(t('报价单', 'Quotation'));
      await b.cell('A1', { value: t('远山科技有限公司', 'Brightline Technologies Ltd.') }); await b.style('A1', { bold: true, size: '14pt', color: C.acc }); await b.merge('A1:D1');
      await b.cell('A2', { value: t('杭州市西湖区文三路 258 号　·　0571-0000 0000　·　sales@example.com', '258 Wensan Road, Hangzhou · +86 571 0000 0000 · sales@example.com') }); await b.style('A2', { color: C.sub, size: '9pt' }); await b.merge('A2:D2');
      await b.cell('E1', { value: t('报 价 单', 'QUOTATION') }); await b.style('E1', { bold: true, size: '20pt', color: C.acc, align: 'right' }); await b.merge('E1:H1');
      await b.values('E2:H4', [[t('报价编号', 'Quote no.'), '', 'QT-20260925-01', ''], [t('报价日期', 'Date'), '', '2026-09-25', ''], [t('有效期至', 'Valid until'), '', '2026-10-25', '']]);
      await b.merge('E2:F2', 'G2:H2', 'E3:F3', 'G3:H3', 'E4:F4', 'G4:H4');
      await b.style('E2:F4', { color: C.sub, align: 'right' }); await b.style('G2:H4', { align: 'right', format: 'yyyy-mm-dd' });
      await b.layout({ heights: { 1: 30 } });
      await b.values('A6:H7', [[t('客户名称', 'Customer'), t('明川文化传媒有限公司', 'Riverstone Media Ltd.'), '', '', t('联系人', 'Contact'), t('张伟', 'Sam Taylor'), '', ''], [t('地址', 'Address'), t('上海市浦东新区世纪大道 100 号', '100 Century Avenue, Shanghai'), '', '', t('电话', 'Phone'), '138 0000 0001', '', '']]);
      await b.merge('B6:D6', 'F6:H6', 'B7:D7', 'F7:H7');
      await b.style('A6:H7', { fill: C.soft, border: 'thin', borderColor: C.line }); await b.style('A6:A7', { bold: true }); await b.style('E6:E7', { bold: true });
      await b.values('A9:H9', [['#', t('品名', 'Item'), t('规格 / 说明', 'Description'), t('单位', 'Unit'), t('数量', 'Qty'), t('单价', 'Unit price'), t('金额', 'Amount'), t('备注', 'Notes')]]);
      await header(b, C, 'A9:H9');
      const items = [
        [t('内容管理系统 · 标准版', 'Content management system, standard'), t('50 个编辑席位，含一年升级', '50 editor seats, one year of upgrades'), t('套', 'licence'), 1, 180000, ''],
        [t('定制开发', 'Custom development'), t('审批流与多渠道发布模块', 'Approval flow and multi-channel publishing'), t('人天', 'day'), 60, 2800, ''],
        [t('数据迁移', 'Data migration'), t('历史内容约 12,000 篇', 'About 12,000 existing articles'), t('项', 'lot'), 1, 36000, ''],
        [t('部署实施', 'Deployment'), t('生产与测试两套环境', 'Production and staging environments'), t('项', 'lot'), 1, 24000, ''],
        [t('培训', 'Training'), t('管理员 1 场，编辑 2 场', 'One admin session, two editor sessions'), t('场', 'session'), 3, 6000, t('可远程', 'Remote available')],
        [t('运维支持', 'Support'), t('工作日 9–18 点，4 小时响应', 'Working days 9–6, four-hour response'), t('月', 'month'), 12, 4500, t('首年', 'First year')],
      ];
      const last = first + items.length - 1;
      await b.values(`A${first}:H${last}`, items.map((x, i) => { const r = first + i; return [i + 1, x[0], x[1], x[2], x[3], x[4], `=E${r}*F${r}`, x[5]]; }));
      const S = last + 1;
      await b.values(`F${S}:G${S + 4}`, [[t('小计', 'Subtotal'), `=SUM(G${first}:G${last})`], [t('折扣', 'Discount'), `=-G${S}*0.05`], [t('折后金额', 'After discount'), `=G${S}+G${S + 1}`], [t('税额（13%）', 'Tax (13%)'), `=ROUND(G${S + 2}*0.13,2)`], [t('合计（含税）', 'Total'), `=G${S + 2}+G${S + 3}`]]);
      await b.style(`F${S}:F${S + 4}`, { align: 'right', color: C.sub });
      await b.style(`F${S + 4}:G${S + 4}`, { bold: true, color: C.acc, size: '12pt' });
      await b.style(`F${first}:G${S + 4}`, { format: CUR });
      await b.style(`A${first}:A${last}`, { align: 'center', color: C.sub });
      await b.style(`D${first}:E${last}`, { align: 'center' });
      await grid(b, C, `A${first}:H${last}`);
      await b.style(`G${S}:G${S + 4}`, { border: 'thin', borderColor: C.line });
      const N = S + 6;
      await b.cell(`A${N}`, { value: t('说明', 'Terms') }); await b.style(`A${N}`, { bold: true, color: C.acc });
      await b.values(`A${N + 1}:A${N + 3}`, [[t('1. 付款方式：合同签订后预付 30%，验收后支付 60%，运维期满支付 10%。', '1. Payment: 30% on signing, 60% on acceptance, 10% at the end of the support period.')], [t('2. 报价含 13% 增值税，开具增值税专用发票。', '2. Prices include 13% VAT; a VAT invoice is issued.')], [t('3. 本报价自报价日起 30 天内有效。', '3. This quotation is valid for 30 days from its date.')]]);
      for (let r = N + 1; r <= N + 3; r++) await b.merge(`A${r}:H${r}`);
      await b.style(`A${N + 1}:H${N + 3}`, { color: C.sub, size: '10pt' });
      await b.values(`A${N + 5}:H${N + 5}`, [[t('报价人：李明', 'Prepared by: Alex Chen'), '', '', '', t('客户确认（签章）：', 'Accepted by (signature):'), '', '', '']]);
      await b.merge(`A${N + 5}:D${N + 5}`, `E${N + 5}:H${N + 5}`);
      await b.layout({ widths: { A: 5, B: 24, C: 30, D: 8, E: 8, F: 13, G: 14, H: 14 }, heights: { 9: 22 } });
    },
  },
];
