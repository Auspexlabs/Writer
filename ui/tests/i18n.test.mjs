// node --test ui/tests/
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const src = readFileSync(new URL('../i18n.js', import.meta.url), 'utf8');

/** Runs i18n.js in a fresh fake page: lang is the stored 设置 › 语言, sys what the desktop app hands over, nav navigator.language. */
function page({ lang, sys, nav = 'en-US', store } = {}) {
  const written = [], html = {};
  const w = {
    localStorage: { getItem: k => k !== 'writer-settings' ? null : store !== undefined ? store : lang ? JSON.stringify({ lang }) : null },
    navigator: { language: nav },
    document: { documentElement: html, currentScript: { src: 'http://127.0.0.1:8813/app/i18n.js' }, write: s => written.push(s) }
  };
  if (sys !== undefined) w.__WRITER_SYS_LANG = sys;
  w.window = w;
  vm.runInNewContext(src, w);
  return Object.assign(w, { written, html });
}

const DICT = {
  '保存': 'Save',
  '共 {n} 页': 'Pages: {n}',
  '个文件': ({ n }) => n === 1 ? '{n} file' : '{n} files',
  '表格@@sheet': 'Sheet',
  '表格': 'Table',
  '共 {{ n }} 页': '{{ n }} pages',
  '说 "你好"': 'Say "hi" <now>'
};

test('Chinese: every call returns the key, placeholders filled, the context dropped, templates untouched', () => {
  const w = page({ lang: '简体中文', sys: 'en-US' });
  w.I18N_EN = DICT;
  assert.equal(w.$lang(), 'zh');
  assert.equal(w.$t('保存'), '保存');
  assert.equal(w.$t('共 {n} 页', { n: 3 }), '共 3 页');
  assert.equal(w.$t('表格@@sheet'), '表格');
  const tpl = '<div title="保存">保存</div>';
  assert.equal(w.$tTemplate(tpl), tpl);
  assert.equal(w.html.lang, 'zh-CN');
  assert.deepEqual(w.written, [], 'no dictionaries in Chinese');
});

test('English: dictionary values, Chinese fallback, vars, plural functions, @@ context', () => {
  const w = page({ lang: 'English' });
  w.I18N_EN = DICT;
  assert.equal(w.$lang(), 'en');
  assert.equal(w.$t('保存'), 'Save');
  assert.equal(w.$t('未翻译'), '未翻译', 'a missing entry falls back to the Chinese');
  assert.equal(w.$t('共 {n} 页', { n: 3 }), 'Pages: 3');
  assert.equal(w.$t('第 {n} 行', { n: 2 }), '第 2 行', 'the fallback gets its vars too');
  assert.equal(w.$t('{a}{b}', { a: 1 }), '1{b}', 'an unknown placeholder stays');
  assert.equal(w.$t('个文件', { n: 1 }), '1 file');
  assert.equal(w.$t('个文件', { n: 5 }), '5 files');
  assert.equal(w.$t('表格@@sheet'), 'Sheet', 'the full key first');
  assert.equal(w.$t('表格@@grid'), 'Table', 'then the bare part');
  assert.equal(w.$t('图表@@chart'), '图表', 'then the Chinese without the context');
  assert.equal(w.html.lang, 'en');
});

test('English: $tTemplate translates exact text and title/placeholder/aria-label/alt values, bindings stay', () => {
  const w = page({ lang: 'English' });
  w.I18N_EN = DICT;
  const T = w.$tTemplate;
  assert.equal(T('<div>\n    保存\n  </div>'), '<div>\n    Save\n  </div>', 'surrounding whitespace kept');
  assert.equal(T('<span>共 {{ n }}\n  页</span>'), '<span>{{ n }} pages</span>', 'bindings inside keys, whitespace runs as one space');
  assert.equal(T('<b>保存文件</b><i>个文件</i>'), '<b>保存文件</b><i>个文件</i>', 'no partial matches, no function values');
  assert.equal(T(`<input placeholder="保存" title='保存' aria-label=" 保存 " alt="保存" data-x="保存" title="{{ tip }}">`),
    `<input placeholder="Save" title='Save' aria-label="Save" alt="Save" data-x="保存" title="{{ tip }}">`);
  assert.equal(T('<p title="说 &quot;你好&quot;">说 "你好"</p>'), '<p title="说 &quot;你好&quot;">Say "hi" &lt;now></p>');
  assert.equal(T('<p title=\'说 "你好"\'>x</p>'), '<p title=\'Say "hi" <now>\'>x</p>');
});

test('follow the system: the desktop app\'s language first, then navigator.language; zh… is Chinese', () => {
  assert.equal(page({ sys: 'zh-Hans-CN', nav: 'en-US' }).$lang(), 'zh', 'the system beats WKWebView\'s navigator.language');
  assert.equal(page({ sys: 'en-GB', nav: 'zh-CN' }).$lang(), 'en');
  assert.equal(page({ sys: '', nav: 'zh-TW' }).$lang(), 'zh', 'no system language: navigator.language');
  assert.equal(page({ nav: 'fr-FR' }).$lang(), 'en');
  assert.equal(page({ lang: '跟随系统', sys: 'ja-JP' }).$lang(), 'en');
  assert.equal(page({ lang: '跟随系统', sys: 'zh-Hant-TW' }).$lang(), 'zh');
  assert.equal(page({ lang: '繁體中文', sys: 'en-US' }).$lang(), 'en', 'an old choice follows the system');
  assert.equal(page({ store: 'not json', sys: 'zh-CN' }).$lang(), 'zh');
  assert.equal(page({ store: 'null', sys: 'en-US' }).$lang(), 'en');
  assert.equal(page({ lang: '简体中文', sys: 'en-US' }).$lang(), 'zh');
  assert.equal(page({ lang: 'English', sys: 'zh-CN' }).$lang(), 'en');
});

test('English loads the dictionaries next to i18n.js; each stub adds to window.I18N_EN', () => {
  const w = page({ lang: 'English' });
  const names = ['chrome', 'shell', 'word', 'slide', 'sheet', 'sheet-fn', 'docs'];
  assert.equal(w.written.join(''), names.map(n => `<script src="http://127.0.0.1:8813/app/i18n/en-${n}.js"></script>`).join(''));
  for (const n of names) vm.runInNewContext(readFileSync(new URL(`../i18n/en-${n}.js`, import.meta.url), 'utf8'), w);
  assert.equal(typeof w.I18N_EN, 'object');
});
