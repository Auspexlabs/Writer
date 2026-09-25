// i18n.js — the UI language. Chinese is the source: every UI string is its own key, and English comes from
// window.I18N_EN, which i18n/en-*.js fill and which loads only in English. A classic script, before support.js on every page.
//   $lang()           'zh' | 'en': 设置 › 语言 (writer-settings.lang) 简体中文 or English, else 跟随系统 (the system's
//                     language; zh… is Chinese, anything else English). Read once: a change applies on the next load.
//   $t(zh, vars)      zh in the current language, {name} placeholders filled from vars. An English value may be a function
//                     (vars) => string, for plurals. '表格@@sheet': Chinese shows 表格; English looks up the full key, then 表格.
//                     A missing English entry falls back to the Chinese.
//   $tTemplate(html)  a template's static text in English (support.js runs every <x-dc> template through it): text between
//                     tags and title / placeholder / aria-label / alt values whose trimmed text, runs of whitespace read as
//                     one space, is a key; {{ }} bindings stay verbatim inside keys ('共 {{ n }} 页'). Unchanged in Chinese.
(function () {
  var w = typeof window !== 'undefined' ? window : globalThis, d = w.document;
  var has = function (o, k) { return !!o && Object.prototype.hasOwnProperty.call(o, k); };
  var stored = '';
  try { stored = JSON.parse(w.localStorage.getItem('writer-settings') || '{}').lang || ''; } catch (e) { }
  // The desktop app hands over the system's preferred language: WKWebView's navigator.language is the app's own localization.
  var sys = w.__WRITER_SYS_LANG || (w.navigator && w.navigator.language) || 'zh';
  var lang = stored === 'English' ? 'en' : stored === '简体中文' ? 'zh' : /^zh/i.test(sys) ? 'zh' : 'en'; // i18n-ok

  var en = function (key) { return lang === 'en' && has(w.I18N_EN, key) ? w.I18N_EN[key] : undefined; };
  var fill = function (s, vars) { return vars ? String(s).replace(/\{(\w+)\}/g, function (m, k) { return has(vars, k) ? vars[k] : m; }) : s; };

  w.$lang = function () { return lang; };

  w.$t = function (zh, vars) {
    var key = String(zh), at = key.indexOf('@@'), bare = at < 0 ? key : key.slice(0, at);
    var v = en(key);
    if (v === undefined && at >= 0) v = en(bare);
    if (v === undefined) v = bare;
    return fill(typeof v === 'function' ? v(vars || {}) : v, vars);
  };

  var tpl = function (text) { var v = en(text.replace(/\s+/g, ' ').trim()); return typeof v === 'string' ? v : undefined; };
  w.$tTemplate = function (html) {
    if (lang !== 'en' || typeof html !== 'string') return html;
    return html
      .replace(/>([^<>]+)</g, function (m, text) {
        var v = tpl(text);
        return v === undefined ? m : '>' + text.replace(/\S(?:[\s\S]*\S)?/, function () { return v.replace(/</g, '&lt;'); }) + '<';
      })
      .replace(/(\s(?:title|placeholder|aria-label|alt)\s*=\s*)(?:"([^"]*)"|'([^']*)')/gi, function (m, pre, dq, sq) {
        var v = tpl(dq !== undefined ? dq : sq);
        if (v === undefined) return m;
        return pre + (dq !== undefined ? '"' + v.replace(/"/g, '&quot;') + '"' : "'" + v.replace(/'/g, '&#39;') + "'");
      });
  };

  if (d && d.documentElement) d.documentElement.lang = lang === 'en' ? 'en' : 'zh-CN';
  // The dictionaries, only in English: parser-inserted, so they run before support.js and the page's own scripts.
  if (d && lang === 'en') {
    var dir = ((d.currentScript && d.currentScript.src) || './i18n.js').replace(/[^/]*$/, 'i18n/');
    d.write(['chrome', 'shell', 'word', 'slide', 'sheet', 'sheet-fn', 'docs'].map(function (n) { return '<script src="' + dir + 'en-' + n + '.js"><\/script>'; }).join(''));
  }
})();
