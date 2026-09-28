// Count actual download-button clicks, independently of the download itself.
// No cookies, identifiers, IP collection, retries or third-party requests.
(() => {
  const count = (event) => {
    if (event.defaultPrevented || (event.type === 'click' ? event.button !== 0 : event.button !== 1)) return;
    const link = event.target.closest?.('a[href]');
    if (!link) return;
    const url = new URL(link.href, location.href);
    if (url.origin !== location.origin) return;
    const platform = { '/download/mac': 'mac', '/download/windows': 'windows' }[url.pathname];
    if (!platform) return;
    const body = JSON.stringify({ platform });
    try {
      if (navigator.sendBeacon?.('/api/download-clicks', body)) return;
    } catch { /* Try keepalive when the beacon cannot be queued. */ }
    try {
      fetch('/api/download-clicks', { method: 'POST', body, keepalive: true, credentials: 'omit' }).catch(() => {});
    } catch { /* A counter failure must never prevent the download. */ }
  };
  document.addEventListener('click', count);
  document.addEventListener('auxclick', count);
})();
