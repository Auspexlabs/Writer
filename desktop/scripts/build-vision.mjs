// Builds writer-vision, the engine's picture helper (Apple Vision background removal, ImageIO re-encoding), as a Tauri
// sidecar: src-tauri/binaries/writer-vision-<target triple>. macOS only; on other hosts it does nothing.
// Run alone for a local build: node desktop/scripts/build-vision.mjs [--universal]
import { spawnSync } from 'node:child_process';
import { mkdirSync, renameSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const desktop = resolve(fileURLToPath(new URL('..', import.meta.url)));
const source = join(desktop, 'vision', 'main.swift');
const binaries = join(desktop, 'src-tauri', 'binaries');
const TRIPLES = { arm64: 'aarch64-apple-darwin', x86_64: 'x86_64-apple-darwin' };

function run(cmd, args) {
  const r = spawnSync(cmd, args, { stdio: 'inherit' });
  if (r.error) throw r.error;
  if (r.status !== 0) throw new Error(`${cmd} ${args[0]} failed`);
}

/** Compiles one architecture into writer-vision-<triple>, renaming the finished binary over the old one (macOS kills a
 *  running process whose code changes under it). */
function compile(arch) {
  const output = join(binaries, `writer-vision-${TRIPLES[arch]}`);
  // macOS 14 is the first with VNGenerateForegroundInstanceMaskRequest, and the app's minimum
  run('xcrun', ['swiftc', '-O', '-target', `${arch}-apple-macos14.0`, source, '-o', output + '.tmp']);
  renameSync(output + '.tmp', output);
  return output;
}

/** Compiles the helper for this Mac's architecture; with `universal`, for arm64 and x86_64 plus one binary joining them.
 *  Returns the sidecar's path, or null on a host that is not macOS. */
export function buildVision({ universal = false } = {}) {
  if (process.platform !== 'darwin') return null;
  mkdirSync(binaries, { recursive: true });
  if (!universal) return compile(process.arch === 'x64' ? 'x86_64' : 'arm64');
  const slices = ['arm64', 'x86_64'].map(compile);
  const output = join(binaries, 'writer-vision-universal-apple-darwin');
  run('lipo', ['-create', ...slices, '-output', output + '.tmp']);
  renameSync(output + '.tmp', output);
  return output;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const output = buildVision({ universal: process.argv.includes('--universal') });
  console.log(output ? `Prepared writer-vision sidecar: ${output}` : 'writer-vision is built on macOS only; skipped.');
}
