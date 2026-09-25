// node scripts/latest-json.mjs <latest.json> <platform> <version> <signature file> <url> [notes]
// Adds one platform's signed update to latest.json, the in-app updater's feed (plugins.updater in tauri.conf.json):
// release-mac.sh web adds darwin-aarch64 and release-win.sh windows-x86_64, in either order. A feed for another version
// starts over, so it never offers this version with another release's files. Empty notes keep the ones already there.
import { readFileSync, writeFileSync } from 'node:fs';

const [file, platform, version, sigFile, url, notes = ''] = process.argv.slice(2);
if (!url) { console.error('usage: node scripts/latest-json.mjs <latest.json> <platform> <version> <signature file> <url> [notes]'); process.exit(2); }
let feed = {};
try { feed = JSON.parse(readFileSync(file, 'utf8')); } catch (e) { }
if (feed.version !== version) feed = {};
const platforms = { ...feed.platforms, [platform]: { signature: readFileSync(sigFile, 'utf8').trim(), url } };
writeFileSync(file, JSON.stringify({ version, notes: notes || feed.notes || '', pub_date: new Date().toISOString(), platforms }, null, 2) + '\n');
