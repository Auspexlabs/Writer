// App Store Connect API for the release setup: signing certificates, the bundle id and the Mac App Store profile.
// Reads a team API key from ASC_KEY_ID, ASC_ISSUER_ID and ASC_KEY_FILE (e.g. `set -a; . <path>/asc_api.env; set +a`).
// The key is only used to sign short-lived request tokens; nothing here prints it.
//
//   node scripts/asc.mjs certs                     list the team's signing certificates
//   node scripts/asc.mjs cert <TYPE>               make a certificate (DEVELOPER_ID_APPLICATION, DISTRIBUTION,
//                                                  MAC_INSTALLER_DISTRIBUTION…) and import it into the login keychain
//   node scripts/asc.mjs fetch <TYPE>              import the newest certificate of a type made elsewhere (e.g. a Developer ID
//                                                  made on the website from a CSR whose key is already in the keychain)
//   node scripts/asc.mjs get <path>                print an API resource (read-only)
//   node scripts/asc.mjs listing <file.json>       store texts, categories, age rating and copyright (store/listing.*.json)
//   node scripts/asc.mjs screenshots <dir> [locale] replace the Mac screenshots of the version being prepared with <dir>/*.png
//   node scripts/asc.mjs free                      price 0 in every territory, available everywhere (new territories too)
//   node scripts/asc.mjs bundle-id                 register cn.thewriter.app for macOS (once)
//   node scripts/asc.mjs profile                   make the Mac App Store profile → ../.signing/Writer_Mac_App_Store.provisionprofile
import { sign, randomBytes, createHash } from 'node:crypto';
import { readFileSync, writeFileSync, mkdtempSync, rmSync, mkdirSync, chmodSync, readdirSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join, resolve, dirname, basename } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';

const BUNDLE_ID = 'cn.thewriter.app';
const repo = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');
const { ASC_KEY_ID: kid, ASC_ISSUER_ID: iss, ASC_KEY_FILE: keyFile } = process.env;
if (!kid || !iss || !keyFile) { console.error('Set ASC_KEY_ID, ASC_ISSUER_ID and ASC_KEY_FILE first.'); process.exit(1); }

function token() {
  const b64 = v => Buffer.from(typeof v === 'string' ? v : JSON.stringify(v)).toString('base64url');
  const now = Math.floor(Date.now() / 1000);
  const body = b64({ alg: 'ES256', kid, typ: 'JWT' }) + '.' + b64({ iss, iat: now, exp: now + 900, aud: 'appstoreconnect-v1' });
  const signature = sign('sha256', Buffer.from(body), { key: readFileSync(keyFile, 'utf8'), dsaEncoding: 'ieee-p1363' });
  return body + '.' + signature.toString('base64url');
}

async function api(method, path, data) {
  const res = await fetch('https://api.appstoreconnect.apple.com' + path, {
    method, headers: { Authorization: 'Bearer ' + token(), 'Content-Type': 'application/json' },
    body: data ? JSON.stringify({ data }) : undefined
  });
  const json = res.status === 204 ? {} : await res.json();
  if (!res.ok) {
    const errors = (json.errors || []).map(e => `${e.code}: ${e.detail || e.title}`).join('\n');
    throw new Error(`${method} ${path} → ${res.status}\n${errors}`);
  }
  return json;
}

async function apiRaw(method, path, body) {
  const res = await fetch('https://api.appstoreconnect.apple.com' + path, {
    method, headers: { Authorization: 'Bearer ' + token(), 'Content-Type': 'application/json' }, body: JSON.stringify(body)
  });
  const json = res.status === 204 ? {} : await res.json();
  if (!res.ok) throw new Error(`${method} ${path} → ${res.status}\n${(json.errors || []).map(e => `${e.code}: ${e.detail || e.title}`).join('\n')}`);
  return json;
}

/** Free everywhere: the $0 price point of the base territory (USA), and every territory available, new ones included. */
async function makeFree() {
  const app = (await api('GET', `/v1/apps?filter[bundleId]=${BUNDLE_ID}`)).data[0];
  const zero = (await api('GET', `/v1/apps/${app.id}/appPricePoints?filter[territory]=USA&limit=5`)).data.find(p => p.attributes.customerPrice === '0.0');
  await apiRaw('POST', '/v1/appPriceSchedules', {
    data: { type: 'appPriceSchedules', relationships: {
      app: { data: { type: 'apps', id: app.id } },
      baseTerritory: { data: { type: 'territories', id: 'USA' } },
      manualPrices: { data: [{ type: 'appPrices', id: '${free}' }] } } },
    included: [{ type: 'appPrices', id: '${free}', attributes: { startDate: null },
      relationships: { appPricePoint: { data: { type: 'appPricePoints', id: zero.id } } } }]
  });
  const territories = (await api('GET', '/v1/territories?limit=200')).data.map(t => t.id);
  const refs = territories.map(t => ({ type: 'territoryAvailabilities', id: '${' + t + '}' }));
  await apiRaw('POST', '/v2/appAvailabilities', {
    data: { type: 'appAvailabilities', attributes: { availableInNewTerritories: true }, relationships: {
      app: { data: { type: 'apps', id: app.id } }, territoryAvailabilities: { data: refs } } },
    included: territories.map(t => ({ type: 'territoryAvailabilities', id: '${' + t + '}',
      attributes: { available: true, releaseDate: null, preOrderEnabled: false },
      relationships: { territory: { data: { type: 'territories', id: t } } } }))
  });
  console.log(`${app.attributes.name}: free (USA base price 0.0), available in ${territories.length} territories and new ones`);
}

const certificates = async () => (await api('GET', '/v1/certificates?limit=200')).data;

async function makeCertificate(type) {
  // key and request stay in a private temporary folder that is removed once the identity is in the keychain
  const dir = mkdtempSync(join(tmpdir(), 'writer-cert-')); chmodSync(dir, 0o700);
  const openssl = '/usr/bin/openssl'; // LibreSSL: its PKCS#12 defaults are the ones `security import` reads
  try {
    execFileSync(openssl, ['req', '-new', '-newkey', 'rsa:2048', '-nodes', '-keyout', join(dir, 'key.pem'),
      '-out', join(dir, 'csr.pem'), '-subj', '/CN=Writer/C=CN'], { stdio: 'ignore' });
    const csr = readFileSync(join(dir, 'csr.pem'), 'utf8');
    const made = await api('POST', '/v1/certificates', { type: 'certificates', attributes: { certificateType: type, csrContent: csr } });
    const a = made.data.attributes;
    writeFileSync(join(dir, 'cert.der'), Buffer.from(a.certificateContent, 'base64'));
    execFileSync(openssl, ['x509', '-inform', 'der', '-in', join(dir, 'cert.der'), '-out', join(dir, 'cert.pem')]);
    const pass = randomBytes(18).toString('base64url');
    execFileSync(openssl, ['pkcs12', '-export', '-inkey', join(dir, 'key.pem'), '-in', join(dir, 'cert.pem'),
      '-out', join(dir, 'id.p12'), '-passout', 'pass:' + pass]);
    execFileSync('security', ['import', join(dir, 'id.p12'), '-P', pass, '-T', '/usr/bin/codesign', '-T', '/usr/bin/productbuild',
      '-T', '/usr/bin/productsign', '-T', '/usr/bin/security'], { stdio: 'inherit' });
    console.log(`${a.certificateType}  ${a.name}  expires ${a.expirationDate}  id ${made.data.id} — in the login keychain`);
    return made.data;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

async function bundleId(create) {
  const found = (await api('GET', `/v1/bundleIds?filter[identifier]=${BUNDLE_ID}&limit=5`)).data.find(b => b.attributes.identifier === BUNDLE_ID);
  if (found || !create) return found;
  const made = await api('POST', '/v1/bundleIds', { type: 'bundleIds', attributes: { identifier: BUNDLE_ID, name: 'Writer', platform: 'MAC_OS' } });
  return made.data;
}


const EDITABLE = ['PREPARE_FOR_SUBMISSION', 'DEVELOPER_REJECTED', 'REJECTED', 'METADATA_REJECTED', 'INVALID_BINARY'];

/** The app, its editable app info and the macOS version being prepared, with their localizations for `locale`. */
async function editable(locale) {
  const app = (await api('GET', `/v1/apps?filter[bundleId]=${BUNDLE_ID}`)).data[0];
  if (!app) throw new Error(`No App Store Connect app for ${BUNDLE_ID}; create it on the website first.`);
  const infos = await api('GET', `/v1/apps/${app.id}/appInfos?include=appInfoLocalizations,ageRatingDeclaration`);
  const info = infos.data.find(i => EDITABLE.includes(i.attributes.state || i.attributes.appStoreState)) || infos.data[0];
  const infoLoc = (infos.included || []).find(x => x.type === 'appInfoLocalizations' && x.attributes.locale === locale);
  const rating = (infos.included || []).find(x => x.type === 'ageRatingDeclarations');
  const versions = await api('GET', `/v1/apps/${app.id}/appStoreVersions?filter[platform]=MAC_OS&include=appStoreVersionLocalizations`);
  const version = versions.data.find(v => EDITABLE.includes(v.attributes.appStoreState || v.attributes.appVersionState));
  if (!version) throw new Error('No macOS version is being prepared.');
  const versionLoc = (versions.included || []).find(x => x.type === 'appStoreVersionLocalizations' && x.attributes.locale === locale
    && version.relationships.appStoreVersionLocalizations.data.some(d => d.id === x.id));
  return { app, info, infoLoc, rating, version, versionLoc };
}

async function applyListing(file) {
  const L = JSON.parse(readFileSync(file, 'utf8'));
  const e = await editable(L.locale);
  const { primaryCategory, secondaryCategory, ...infoTexts } = L.appInfo;
  await api('PATCH', `/v1/appInfoLocalizations/${e.infoLoc.id}`, { type: 'appInfoLocalizations', id: e.infoLoc.id, attributes: infoTexts });
  await api('PATCH', `/v1/appInfos/${e.info.id}`, { type: 'appInfos', id: e.info.id, relationships: {
    primaryCategory: { data: { type: 'appCategories', id: primaryCategory } },
    secondaryCategory: { data: secondaryCategory ? { type: 'appCategories', id: secondaryCategory } : null } } });
  await api('PATCH', `/v1/ageRatingDeclarations/${e.rating.id}`, { type: 'ageRatingDeclarations', id: e.rating.id, attributes: L.ageRating });
  const { copyright, ...versionTexts } = L.version;
  await api('PATCH', `/v1/appStoreVersions/${e.version.id}`, { type: 'appStoreVersions', id: e.version.id, attributes: { copyright } });
  await api('PATCH', `/v1/appStoreVersionLocalizations/${e.versionLoc.id}`, { type: 'appStoreVersionLocalizations', id: e.versionLoc.id, attributes: versionTexts });
  console.log(`${e.app.attributes.name} ${e.version.attributes.versionString} (${L.locale}): texts, categories, age rating and copyright updated`);
}

async function uploadScreenshots(dir, locale) {
  const files = readdirSync(dir).filter(f => f.toLowerCase().endsWith('.png')).sort().map(f => join(dir, f));
  if (!files.length) throw new Error(`No .png files in ${dir}`);
  const e = await editable(locale);
  const sets = await api('GET', `/v1/appStoreVersionLocalizations/${e.versionLoc.id}/appScreenshotSets`);
  let set = sets.data.find(s => s.attributes.screenshotDisplayType === 'APP_DESKTOP');
  if (!set) set = (await api('POST', '/v1/appScreenshotSets', { type: 'appScreenshotSets', attributes: { screenshotDisplayType: 'APP_DESKTOP' },
    relationships: { appStoreVersionLocalization: { data: { type: 'appStoreVersionLocalizations', id: e.versionLoc.id } } } })).data;
  for (const old of (await api('GET', `/v1/appScreenshotSets/${set.id}/appScreenshots?limit=50`)).data) await api('DELETE', `/v1/appScreenshots/${old.id}`);
  const ids = [];
  for (const file of files) {
    const bytes = readFileSync(file);
    const made = (await api('POST', '/v1/appScreenshots', { type: 'appScreenshots', attributes: { fileName: basename(file), fileSize: bytes.length },
      relationships: { appScreenshotSet: { data: { type: 'appScreenshotSets', id: set.id } } } })).data;
    for (const op of made.attributes.uploadOperations) {
      const headers = Object.fromEntries((op.requestHeaders || []).map(h => [h.name, h.value]));
      const res = await fetch(op.url, { method: op.method, headers, body: bytes.subarray(op.offset, op.offset + op.length) });
      if (!res.ok) throw new Error(`upload of ${basename(file)} failed: ${res.status}`);
    }
    await api('PATCH', `/v1/appScreenshots/${made.id}`, { type: 'appScreenshots', id: made.id,
      attributes: { uploaded: true, sourceFileChecksum: createHash('md5').update(bytes).digest('hex') } });
    ids.push(made.id);
    console.log(`uploaded ${basename(file)}`);
  }
  await api('PATCH', `/v1/appScreenshotSets/${set.id}/relationships/appScreenshots`, ids.map(id => ({ type: 'appScreenshots', id })));
  for (let i = 0; i < 30; i++) { // Apple processes each image; report what it thinks of them
    const states = (await api('GET', `/v1/appScreenshotSets/${set.id}/appScreenshots?limit=50`)).data.map(s => [s.attributes.fileName, s.attributes.assetDeliveryState?.state, s.attributes.assetDeliveryState?.errors]);
    if (states.every(([, st]) => st === 'COMPLETE' || st === 'FAILED')) { for (const s of states) console.log(s[0], s[1], s[2]?.length ? JSON.stringify(s[2]) : ''); return; }
    await new Promise(r => setTimeout(r, 4000));
  }
  console.log('Apple is still processing the screenshots; check App Store Connect in a minute.');
}

const [command, arg] = process.argv.slice(2);
switch (command) {
  case 'certs':
    for (const c of await certificates()) console.log(`${c.attributes.certificateType.padEnd(28)} ${c.attributes.name}  expires ${c.attributes.expirationDate}  id ${c.id}`);
    break;
  case 'cert':
    if (!arg) { console.error('cert <TYPE>'); process.exit(1); }
    await makeCertificate(arg);
    break;
  case 'fetch': {
    const newest = (await certificates()).filter(c => c.attributes.certificateType === arg)
      .sort((x, y) => y.attributes.expirationDate.localeCompare(x.attributes.expirationDate))[0];
    if (!newest) { console.error(`No ${arg} certificate on the team.`); process.exit(1); }
    const dir = mkdtempSync(join(tmpdir(), 'writer-cert-'));
    try {
      writeFileSync(join(dir, 'cert.cer'), Buffer.from(newest.attributes.certificateContent, 'base64'));
      execFileSync('security', ['import', join(dir, 'cert.cer'), '-k', join(process.env.HOME, 'Library/Keychains/login.keychain-db')], { stdio: 'inherit' });
    } finally { rmSync(dir, { recursive: true, force: true }); }
    console.log(`${arg}  ${newest.attributes.name}  expires ${newest.attributes.expirationDate}  id ${newest.id} — imported`);
    break;
  }
  case 'free':
    await makeFree();
    break;
  case 'listing':
    await applyListing(arg);
    break;
  case 'screenshots':
    await uploadScreenshots(arg, process.argv[4] || 'zh-Hans');
    break;
  case 'get': // read-only: print any API resource, e.g. get /v1/apps/<id>/appInfos
    console.log(JSON.stringify(await api('GET', arg), null, 2));
    break;
  case 'bundle-id': {
    const b = await bundleId(true);
    console.log(`${b.attributes.identifier}  ${b.attributes.platform}  id ${b.id}`);
    break;
  }
  case 'profile': {
    const b = await bundleId(false);
    if (!b) { console.error('Register the bundle id first: node scripts/asc.mjs bundle-id'); process.exit(1); }
    const dist = (await certificates()).filter(c => ['DISTRIBUTION', 'MAC_APP_DISTRIBUTION'].includes(c.attributes.certificateType))
      .sort((x, y) => y.attributes.expirationDate.localeCompare(x.attributes.expirationDate));
    if (!dist.length) { console.error('Make a distribution certificate first: node scripts/asc.mjs cert DISTRIBUTION'); process.exit(1); }
    const made = await api('POST', '/v1/profiles', {
      type: 'profiles', attributes: { name: `Writer Mac App Store ${new Date().toISOString().slice(0, 10)}`, profileType: 'MAC_APP_STORE' },
      relationships: { bundleId: { data: { type: 'bundleIds', id: b.id } }, certificates: { data: dist.map(c => ({ type: 'certificates', id: c.id })) } }
    });
    const out = join(repo, '.signing', 'Writer_Mac_App_Store.provisionprofile');
    mkdirSync(dirname(out), { recursive: true });
    writeFileSync(out, Buffer.from(made.data.attributes.profileContent, 'base64'));
    console.log(`${made.data.attributes.name}  expires ${made.data.attributes.expirationDate} → ${out}`);
    break;
  }
  default:
    console.log(readFileSync(fileURLToPath(import.meta.url), 'utf8').split('\n').slice(0, 9).join('\n'));
}
