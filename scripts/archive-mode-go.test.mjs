import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { inspectZipArchive } from '../packages/pptx-parser/dist/index.js';
import { zip } from '../packages/pptx-parser/test/zip-fixture.mjs';

test('OS X ZIP file type checks agree with Go mode interpretation', async t => {
  try { execFileSync('go', ['version'], { stdio: 'pipe' }); } catch { t.skip('Go is unavailable'); return; }
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-zip-modes-')); t.after(() => rm(dir, { recursive: true, force: true }));
  const source = path.join(dir, 'modes.go');
  await writeFile(source, `package main
import ("archive/zip"; "encoding/json"; "os")
func main() { var rows []string; for _, name := range os.Args[1:] { r, err := zip.OpenReader(name); if err != nil { panic(err) }; mode := r.File[0].Mode(); switch { case mode&os.ModeSymlink != 0: rows = append(rows, "link"); case mode.IsDir(): rows = append(rows, "directory"); case mode.IsRegular(): rows = append(rows, "regular"); default: rows = append(rows, "special") }; r.Close() }; json.NewEncoder(os.Stdout).Encode(rows) }
`);
  const modes = [0x8000, 0x4000, 0xa000, 0x1000], archives = [];
  for (const [index, mode] of modes.entries()) {
    const { bytes, directory } = zip([{ name: mode === 0x4000 ? 'folder/' : 'file', method: 0 }]);
    bytes.writeUInt16LE((19 << 8) | 20, directory + 4); bytes.writeUInt32LE((mode << 16) >>> 0, directory + 38);
    const file = path.join(dir, `${index}.zip`); await writeFile(file, bytes); archives.push({ file, bytes });
  }
  const observed = JSON.parse(execFileSync('go', ['run', source, ...archives.map(item => item.file)], { encoding: 'utf8', stdio: 'pipe' }));
  assert.deepEqual(observed, ['regular', 'directory', 'link', 'special']);
  assert.equal((await inspectZipArchive(archives[0].bytes)).entries[0].directory, false);
  assert.equal((await inspectZipArchive(archives[1].bytes)).entries[0].directory, true);
  for (const item of archives.slice(2)) await assert.rejects(inspectZipArchive(item.bytes), { code: 'archive-unsupported' });
});
