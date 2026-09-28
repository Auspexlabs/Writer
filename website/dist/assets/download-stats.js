const number = new Intl.NumberFormat('zh-CN');
const time = new Intl.DateTimeFormat('zh-CN', { timeZone: 'Asia/Shanghai', dateStyle: 'short', timeStyle: 'medium', hour12: false });
const status = document.getElementById('status');
const refresh = document.getElementById('refresh');
let loading = false;
async function load() {
  if (loading) return;
  loading = true;
  refresh.disabled = true;
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 10000);
  try {
    const response = await fetch('/stats/data.json', { cache: 'no-cache', signal: controller.signal });
    if (!response.ok) throw new Error('unavailable');
    const data = await response.json();
    document.querySelectorAll('[data-stat]').forEach((el) => { el.textContent = number.format(data[el.dataset.stat]); });
    document.getElementById('since').textContent = '自 ' + time.format(new Date(data.started_at));
    const rows = data.daily.filter((day) => day.day >= data.started_at.slice(0, 10)).map((day) => {
      const row = document.createElement('tr');
      [day.day, number.format(day.mac), number.format(day.windows), number.format(day.total)].forEach((value) => {
        const cell = document.createElement('td');
        cell.textContent = value;
        row.append(cell);
      });
      return row;
    });
    document.getElementById('days').replaceChildren(...rows);
    document.getElementById('empty').hidden = data.total !== 0;
    status.textContent = '更新于 ' + time.format(new Date(data.updated_at)) + ' · 每 30 秒自动刷新';
  } catch {
    status.textContent = '暂时无法读取最新统计，请稍后刷新。已有数字为上次读取结果。';
  } finally {
    clearTimeout(timeout);
    loading = false;
    refresh.disabled = false;
  }
}
refresh.addEventListener('click', load);
setInterval(() => { if (!document.hidden) load(); }, 30000);
document.addEventListener('visibilitychange', () => { if (!document.hidden) load(); });
load();
