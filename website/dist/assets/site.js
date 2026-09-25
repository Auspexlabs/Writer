// The current version next to the download buttons, from download/latest.json ({"version","size","date"}; size is the dmg's, not shown).
// When the file cannot be read the line stays hidden; the buttons work either way.
fetch('download/latest.json', { cache: 'no-cache' })
  .then((r) => (r.ok ? r.json() : Promise.reject(r.status)))
  .then(({ version, size, date }) => {
    if (typeof version !== 'string' || !version) return;
    let text = '版本 ' + version;
    if (typeof date === 'string' && date) text += '，' + date + ' 更新';
    document.querySelectorAll('[data-latest]').forEach((el) => { el.textContent = text; el.hidden = false; });
  })
  .catch(() => {});

// The hero's download button follows the visitor's system; the download section offers both.
if (/Windows/i.test(navigator.userAgent)) {
  const b = document.querySelector('[data-dl-hero]');
  if (b) { b.href = 'download/windows'; b.textContent = '下载 Windows 版'; }
}

// Sections rise into place once. <head> set the "js" class that hides them, and the fallback there waits for this flag.
window.__reveal = true;
const reveal = new IntersectionObserver((entries) => {
  entries.forEach((e) => { if (e.isIntersecting) { e.target.classList.add('in'); reveal.unobserve(e.target); } });
}, { rootMargin: '0px 0px -8% 0px', threshold: 0.12 });
document.querySelectorAll('.reveal').forEach((el) => reveal.observe(el));

// Editor tabs: one large window at a time. Without this script every panel simply shows.
const tabs = [...document.querySelectorAll('[role="tab"]')];
const select = (tab, focus) => {
  tabs.forEach((t) => {
    const on = t === tab;
    t.setAttribute('aria-selected', String(on));
    t.tabIndex = on ? 0 : -1;
    document.getElementById(t.getAttribute('aria-controls')).hidden = !on;
  });
  if (focus) tab.focus();
};
tabs.forEach((t, i) => {
  t.addEventListener('click', () => select(t));
  t.addEventListener('keydown', (e) => {
    const step = { ArrowRight: 1, ArrowLeft: -1 }[e.key];
    if (step) { e.preventDefault(); select(tabs[(i + step + tabs.length) % tabs.length], true); }
  });
});
if (tabs.length) select(tabs.find((t) => t.getAttribute('aria-selected') === 'true') || tabs[0]);

// The Mac-details rail: the arrow buttons page by one card; swiping and the keyboard work on their own.
const rail = document.querySelector('.rail');
document.querySelectorAll('[data-rail]').forEach((button) => {
  button.addEventListener('click', () => {
    const card = rail.querySelector('.card');
    rail.scrollBy({ left: Number(button.dataset.rail) * (card.offsetWidth + 20), behavior: 'smooth' });
  });
});

document.querySelectorAll('[data-copy]').forEach((button) => {
  button.addEventListener('click', async () => {
    const code = document.getElementById(button.dataset.copy);
    try {
      await navigator.clipboard.writeText(code.textContent.trim());
      button.textContent = '已复制';
    } catch {
      button.textContent = '复制失败';
    }
    setTimeout(() => { button.textContent = '复制'; }, 1800);
  });
});
