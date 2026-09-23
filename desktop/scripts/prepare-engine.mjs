// Builds the `writer` engine sidecar into src-tauri/binaries/writer-<target triple>.
//
//   node scripts/prepare-engine.mjs            development: a self-contained single-file build for this machine
//   node scripts/prepare-engine.mjs --release  distribution: NativeAOT (no JIT, nothing unpacked at run time, so it signs,
//                                              notarizes and runs in the App Store sandbox as-is); on macOS a universal
//                                              arm64 + x86_64 binary for `tauri build --target universal-apple-darwin`
//
// Release builds need Microsoft's .NET SDK. Homebrew's build links /opt/homebrew's brotli (and OpenSSL for NativeAOT), so
// its engine does not even start on a Mac without Homebrew, nor inside the App Store sandbox (dyld: libbrotlidec not
// loaded). Set DOTNET to the SDK's dotnet, or install it into ~/.dotnet.
import { spawnSync } from 'node:child_process';
import { mkdirSync, copyFileSync, existsSync, rmSync, renameSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { homedir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { buildVision } from './build-vision.mjs';

const desktop = resolve(fileURLToPath(new URL('..', import.meta.url)));
const root = resolve(desktop, '..');
const release = process.argv.includes('--release');
const targets = {
  'darwin-arm64': ['osx-arm64', 'aarch64-apple-darwin'],
  'darwin-x64': ['osx-x64', 'x86_64-apple-darwin'],
  'win32-x64': ['win-x64', 'x86_64-pc-windows-msvc'],
  'win32-arm64': ['win-arm64', 'aarch64-pc-windows-msvc']
};
const host = targets[`${process.platform}-${process.arch}`];
if (!host) throw new Error(`Unsupported desktop target: ${process.platform}-${process.arch}`);

const localSdk = join(homedir(), '.dotnet', process.platform === 'win32' ? 'dotnet.exe' : 'dotnet');
const dotnet = process.env.DOTNET || (existsSync(localSdk) ? localSdk : 'dotnet');

function run(cmd, args) {
  const result = spawnSync(cmd, args, { cwd: root, stdio: 'inherit' });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

if (release) {
  const sdks = spawnSync(dotnet, ['--list-sdks'], { encoding: 'utf8' }).stdout || '';
  if (/homebrew|\/Cellar\//i.test(sdks)) {
    console.error(`Release builds need Microsoft's .NET SDK, not Homebrew's (${dotnet} lists: ${sdks.trim()}).\n` +
      `Install it with: curl -sSL https://dot.net/v1/dotnet-install.sh | bash -s -- --channel 10.0   (goes to ~/.dotnet)`);
    process.exit(1);
  }
}

/** Publishes the engine for one runtime and returns the path of the `writer` executable. */
function publish(rid) {
  const output = join(desktop, '.engine', rid);
  rmSync(output, { recursive: true, force: true });
  mkdirSync(output, { recursive: true });
  const mode = release
    ? ['-p:PublishAot=true', '-p:StripSymbols=true']
    : ['-p:PublishSingleFile=true', '-p:PublishTrimmed=true', '-p:EnableCompressionInSingleFile=true', '-p:IncludeNativeLibrariesForSelfExtract=true'];
  run(dotnet, ['publish', 'src/Writer.Cli/Writer.Cli.csproj', '-c', 'Release', '-r', rid, '--self-contained', ...mode,
    '-p:DebugType=none', '-o', output, '--nologo', '-v', 'quiet']);
  return join(output, process.platform === 'win32' ? 'writer.exe' : 'writer');
}

/** Replaces a binary by renaming a finished copy over it: a copy in place would change the pages of a running engine,
 *  and macOS kills a process whose code no longer matches its signature. */
function install(source, destination) {
  copyFileSync(source, destination + '.tmp');
  renameSync(destination + '.tmp', destination);
}

const binaries = join(desktop, 'src-tauri', 'binaries');
mkdirSync(binaries, { recursive: true });
const extension = process.platform === 'win32' ? '.exe' : '';

if (release && process.platform === 'darwin') {
  // Both architectures, then one universal binary; the per-arch copies serve single-architecture builds.
  const arm = publish('osx-arm64'), x64 = publish('osx-x64');
  install(arm, join(binaries, 'writer-aarch64-apple-darwin'));
  install(x64, join(binaries, 'writer-x86_64-apple-darwin'));
  const universal = join(binaries, 'writer-universal-apple-darwin');
  run('lipo', ['-create', arm, x64, '-output', universal + '.tmp']);
  renameSync(universal + '.tmp', universal);
  console.log(`Prepared Writer sidecar (NativeAOT, universal): ${join(binaries, 'writer-universal-apple-darwin')}`);
} else {
  const [rid, triple] = host;
  const destination = join(binaries, `writer-${triple}${extension}`);
  install(publish(rid), destination);
  console.log(`Prepared Writer sidecar${release ? ' (NativeAOT)' : ''}: ${destination}`);
}

// the picture helper (抠图, compression) ships beside the engine on macOS; nothing to do elsewhere
const vision = buildVision({ universal: release && process.platform === 'darwin' });
if (vision) console.log(`Prepared writer-vision: ${vision}`);
