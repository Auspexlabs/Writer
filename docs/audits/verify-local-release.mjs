// Read-only validation of the local Mac release, bundled UI and updater feeds.
// Run from the repository root after desktop/scripts/release-mac.sh web.
import {readFileSync,readdirSync,statSync} from 'node:fs';
import {join,relative,resolve} from 'node:path';
import {createHash,createPublicKey,verify} from 'node:crypto';
import {execFileSync} from 'node:child_process';
import assert from 'node:assert/strict';
const root=process.cwd(),config=JSON.parse(readFileSync('desktop/src-tauri/tauri.conf.json','utf8')),version=config.version;
const app=resolve(process.argv[2]||'desktop/src-tauri/target/universal-apple-darwin/release/bundle/macos/Writer.app');
const files=directory=>readdirSync(directory,{withFileTypes:true}).flatMap(e=>e.isDirectory()?files(join(directory,e.name)):e.isFile()?[join(directory,e.name)]:[]);
const hash=file=>createHash('sha256').update(readFileSync(file)).digest('hex');
const source=join(root,'ui'),bundled=join(app,'Contents/Resources/ui'),sourceFiles=files(source).map(p=>relative(source,p)).sort(),bundleFiles=files(bundled).map(p=>relative(bundled,p)).sort();
assert.deepEqual(bundleFiles,sourceFiles,'Bundled UI file inventory differs from source');
for(const file of sourceFiles)assert.equal(hash(join(source,file)),hash(join(bundled,file)),file);
const prefix=`Writer-${version}-mac`,dmg=join(root,'desktop/dist',prefix+'.dmg'),archive=join(root,'desktop/dist',prefix+'.app.tar.gz');
const signature=readFileSync(archive+'.sig','utf8').trim(),lines=Buffer.from(signature,'base64').toString('utf8').trimEnd().split(/\r?\n/);
const keyLines=Buffer.from(config.plugins.updater.pubkey,'base64').toString('utf8').trim().split(/\r?\n/),rawKey=Buffer.from(keyLines[1],'base64'),rawSig=Buffer.from(lines[1],'base64');
assert.equal(rawKey.length,42);assert.equal(rawSig.length,74);assert.deepEqual(rawKey.subarray(2,10),rawSig.subarray(2,10),'Updater key id differs');
assert.ok(lines[2].startsWith('trusted comment: '));assert.ok(['Ed','ED'].includes(rawSig.subarray(0,2).toString()));
const publicKey=createPublicKey({key:Buffer.concat([Buffer.from('302a300506032b6570032100','hex'),rawKey.subarray(10)]),format:'der',type:'spki'});
const data=readFileSync(archive),signed=rawSig[1]===68?createHash('blake2b512').update(data).digest():data;
assert.ok(verify(null,signed,publicKey,rawSig.subarray(10)),'Updater archive signature is invalid');
assert.ok(verify(null,Buffer.concat([rawSig.subarray(10),Buffer.from(lines[2].slice(17))]),publicKey,Buffer.from(lines[3],'base64')),'Updater trusted comment signature is invalid');
for(const location of ['desktop/dist/latest.json','desktop/dist/updates/mac/latest.json']){
 const feed=JSON.parse(readFileSync(location,'utf8'));assert.equal(feed.version,version);
 for(const platform of ['darwin-aarch64','darwin-x86_64']){assert.equal(feed.platforms[platform].signature,signature);assert.ok(feed.platforms[platform].url.endsWith(prefix+'.app.tar.gz'));}
}
assert.equal(hash('desktop/dist/updates/mac/'+prefix+'.app.tar.gz'),hash(archive));
assert.equal(readFileSync('desktop/dist/updates/mac/'+prefix+'.app.tar.gz.sig','utf8').trim(),signature);
const plist=key=>execFileSync('/usr/libexec/PlistBuddy',['-c','Print :'+key,join(app,'Contents/Info.plist')],{encoding:'utf8'}).trim();
assert.equal(plist('CFBundleShortVersionString'),version);
console.log(JSON.stringify({generatedAt:new Date().toISOString(),appVersion:version,appBuild:plist('CFBundleVersion'),matchedUiFiles:sourceFiles.length,dmgBytes:statSync(dmg).size,dmgSha256:hash(dmg),appBytes:files(app).reduce((n,p)=>n+statSync(p).size,0),updaterArchiveBytes:data.length,updaterSignatureAndFeedVerified:true,scope:'Bundle UI matches current source; archive Ed25519, trusted comment, both feeds and duplicate archive verified. Not a publication or installation check.'},null,2));
