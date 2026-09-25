// Runs at document start in every Writer window, before the page's own scripts (Tauri initialization script).
// main.rs prepends: window.__WRITER_NATIVE__ = { kind: 'main' | 'settings' | 'about' | 'gestures', store: { key: value } }
(() => {
  const N = window.__WRITER_NATIVE__, T = window.__TAURI_INTERNALS__;
  if (!N || !T || location.protocol !== 'http:') return;
  const invoke = (cmd, args) => T.invoke(cmd, args).catch(() => {});
  const ls = window.localStorage, setItem = Storage.prototype.setItem;

  // Shared settings. Every window's engine listens on its own port, so each window (and each launch) is a new origin
  // with an empty localStorage: Rust keeps these keys, seeds them once per window and relays every change
  // (settings_set → the other windows' __writerStore). Same-origin windows also get WebKit's own storage event.
  const KEYS = ['writer-settings', 'writer-mac'], sent = {};
  try {
    if (sessionStorage.getItem('writer-native-seeded') !== '1') {
      for (const k in N.store) setItem.call(ls, k, N.store[k]);
      sessionStorage.setItem('writer-native-seeded', '1');
    }
    KEYS.forEach(k => { sent[k] = ls.getItem(k); });
  } catch (e) { }
  const push = k => {
    const v = ls.getItem(k);
    if (v !== null && v !== sent[k]) { sent[k] = v; invoke('settings_set', { key: k, json: v }); }
  };
  Storage.prototype.setItem = function (k, v) {
    setItem.call(this, k, v);
    if (this === ls && KEYS.includes(k)) push(k);
  };
  window.addEventListener('writer-settings', () => KEYS.forEach(push));
  window.addEventListener('storage', e => { if (KEYS.includes(e.key)) push(e.key); });
  const theme = () => {
    let t = 'light';
    try { t = JSON.parse(ls.getItem('writer-settings') || '{}').theme || 'light'; } catch (e) { }
    if (t === 'auto') t = matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    document.documentElement.dataset.theme = t;
  };
  theme();
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', theme);
  let remount = () => { };
  window.__writerStore = (k, v) => { // another window changed a shared key
    sent[k] = v;
    setItem.call(ls, k, v);
    theme();
    window.dispatchEvent(new Event('writer-settings'));
    remount();
  };

  if (N.kind === 'main') {
    // ⌘O belongs to the native 文件 › 打开… panel, which opens the file in place; the shell's own ⌘O uploads a copy.
    // Stopping the event here (without preventDefault) leaves it unhandled, so WebKit hands it to the menu bar.
    window.addEventListener('keydown', e => {
      if (e.metaKey && !e.ctrlKey && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'o') e.stopImmediatePropagation();
    }, true);
    return;
  }

  // Settings / about / gestures: the design's glass pages in a transparent window over macOS vibrancy.
  const name = decodeURIComponent(location.pathname.split('/').pop().replace(/\.dc\.html$/, ''));
  const props = { onClose: () => invoke('aux_close') };
  const tab = new URLSearchParams(location.search).get('tab');
  if (tab) props.tab = tab;
  let n = 0;
  const setProps = extra => window.__dcSetProps && window.__dcSetProps(name, Object.assign({}, props, extra));
  (function boot() { if (window.__dcSetProps) setProps(); else setTimeout(boot, 20); })();
  remount = () => setProps({ key: 'r' + ++n }); // the settings page reads localStorage once, when it mounts
  window.__writerTab = t => { props.tab = t; remount(); };

  const css = document.createElement('style');
  // Over vibrancy the window is transparent and the page adds a light tint; the App Store build has no transparent
  // windows (private API), so there the page paints the same glass colours on an opaque window.
  css.textContent = (N.opaque ? `
html,body{background:var(--nglass)!important}
html{--nglass:linear-gradient(180deg,#F7F7F9,#EFEFF2)}
html[data-theme="dark"]{--nglass:linear-gradient(180deg,#2C2C2F,#1E1E20)}` : `
html,body{background:transparent!important}
html{--nglass:linear-gradient(180deg,rgba(255,255,255,.46),rgba(246,246,248,.3))}
html[data-theme="dark"]{--nglass:linear-gradient(180deg,rgba(50,50,54,.46),rgba(30,30,32,.36))}`) + `
#dc-root>.sc-host>div:first-child{width:100vw!important;min-height:100vh;border-radius:0!important;background:var(--nglass)!important;
  backdrop-filter:none!important;-webkit-backdrop-filter:none!important;box-shadow:none!important}
#dc-root>.sc-host[data-sc-name="MacSettings"]>div:first-child{height:100vh;display:flex;flex-direction:column}
#dc-root>.sc-host[data-sc-name="MacSettings"]>div:first-child>div:nth-child(2){flex:1;min-height:0;max-height:none!important}
button[data-close]{position:relative}
button[data-close]::before{content:'';position:absolute;inset:-14px}`;
  document.documentElement.appendChild(css);

  // The page's own lights stand in for the hidden native ones (✕ → onClose); its header strip drags the window, Esc closes it.
  // The lights are drawn at 12px, well under the header's own height, so a mousedown that is a few pixels off the
  // button (still meant for it) used to land on the header background and start a drag instead of closing the window;
  // the ::before above gives the close button a bigger hit box so those clicks land on the button itself.
  document.addEventListener('mousedown', e => {
    if (e.button === 0 && e.detail === 1 && e.target.matches && e.target.matches('#dc-root>.sc-host>div>div:first-child')) {
      e.preventDefault();
      invoke('plugin:window|start_dragging');
    }
  });
  window.addEventListener('keydown', e => { if (e.key === 'Escape' && !e.defaultPrevented) invoke('aux_close'); });
})();
