import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

async function setup(t) {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const source = await readFile('packages/svg-renderer/dist/browser/media.js', 'utf8');
  await page.evaluate(async code => { const url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' })); try { window.adapter = await import(url); } finally { URL.revokeObjectURL(url); } }, source);
  const archive = readGeneratedZip(await readFile('tests/fixtures/powerpoint-caption-playback.pptx'));
  const movie = [...archive.entries()].find(([name]) => name.endsWith('.mp4'))[1];
  return { page, requests, movie: [...movie] };
}

test('explicit blob mount plays real media, enforces trim and releases each instance', async t => {
  const { page, requests, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const wrap = document.createElement('div'); document.body.append(wrap);
    const originalCreate = URL.createObjectURL, originalRevoke = URL.revokeObjectURL, live = new Set();
    URL.createObjectURL = value => { const u = originalCreate.call(URL, value); live.add(u); return u; };
    URL.revokeObjectURL = u => { live.delete(u); originalRevoke.call(URL, u); };
    const source = new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' });
    const a = adapter.mountBrowserMedia(wrap, { kind: 'video', source, contentType: 'video/mp4', trimStartMs: 2000, trimEndMs: 2000, muted: true, rewindAfterPlayback: true });
    const b = adapter.mountBrowserMedia(wrap, { kind: 'video', source, contentType: 'video/mp4', muted: true });
    const seen = [], unsubscribe = a.subscribe(s => seen.push(s.status));
    const ready = x => new Promise((resolve, reject) => { const id = setTimeout(() => reject(Error('metadata timeout')), 5000); const off = x.subscribe(s => { if (s.status === 'ready') { clearTimeout(id); queueMicrotask(off); resolve(); } else if (s.status === 'error') { clearTimeout(id); reject(Error(s.failure)); } }); });
    await Promise.all([ready(a), ready(b)]);
    const initial = { time: a.element.currentTime, duration: a.state.durationMs, capability: a.state.capability, live: live.size };
    await a.play(); await new Promise(resolve => setTimeout(resolve, 120));
    const playing = !a.element.paused && b.element.paused;
    a.pause(); await new Promise(resolve => setTimeout(resolve, 30));
    a.seek(9500); a.element.dispatchEvent(new Event('timeupdate'));
    a.seek(10000); a.element.dispatchEvent(new Event('timeupdate'));
    const ended = { status: a.state.status, time: a.element.currentTime };
    a.dispose(); a.dispose(); const afterFirst = { live: live.size, children: wrap.children.length, state: a.state.status };
    unsubscribe(); b.dispose();
    const afterAll = { live: live.size, children: wrap.children.length, state: b.state.status };
    URL.createObjectURL = originalCreate; URL.revokeObjectURL = originalRevoke;
    return { initial, playing, ended, afterFirst, afterAll, seen };
  }, movie);
  assert.equal(result.initial.time, 2); assert.ok(Math.abs(result.initial.duration - 12000) < 30); assert.ok(['probably', 'maybe'].includes(result.initial.capability)); assert.equal(result.initial.live, 2);
  assert.equal(result.playing, true); assert.deepEqual(result.ended, { status: 'ended', time: 2 });
  assert.deepEqual(result.afterFirst, { live: 1, children: 1, state: 'disposed' }); assert.deepEqual(result.afterAll, { live: 0, children: 0, state: 'disposed' });
  assert.ok(result.seen.includes('playing')); assert.ok(result.seen.includes('ended')); assert.equal(result.seen.at(-1), 'disposed'); assert.deepEqual(requests, []);
});

test('capability is advisory, permission rejection can retry, and late failures cannot revive disposal', async t => {
  const { page, requests, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const source = new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' });
    const c = adapter.mountBrowserMedia(host, { kind: 'video', source, contentType: 'video/not-a-codec', muted: true });
    const capability = c.state.capability, realPlay = c.element.play.bind(c.element);
    c.element.play = () => Promise.reject(new DOMException('gesture', 'NotAllowedError'));
    let failure; try { await c.play(); } catch { failure = c.state.failure; }
    c.element.play = realPlay; await c.play(); const recovered = c.state.failure === undefined;
    let rejectLater; c.element.play = () => new Promise((_, reject) => { rejectLater = reject; });
    const pending = c.play().catch(() => {}); c.dispose(); rejectLater(new DOMException('late', 'NotSupportedError')); await pending;
    let disposedError = false; try { c.pause(); } catch { disposedError = true; }
    return { capability, failure, recovered, status: c.state.status, disposedError, children: host.children.length };
  }, movie);
  assert.deepEqual(result, { capability: 'unsupported', failure: 'permission', recovered: true, status: 'disposed', disposedError: true, children: 0 }); assert.deepEqual(requests, []);
});

test('invalid inputs and invalid retained duration fail explicitly without leaked nodes', async t => {
  const { page, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host); const source = new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' });
    const errors = [];
    for (const opts of [{ kind: 'video', source, contentType: 'video/mp4', volume: 2 }, { kind: 'video', source, contentType: 'video/mp4', loop: 1 }, { kind: 'video', source, contentType: 'video/mp4', trimStartMs: -1 }]) { try { adapter.mountBrowserMedia(host, opts); } catch (e) { errors.push(e.name); } }
    const c = adapter.mountBrowserMedia(host, { kind: 'video', source, contentType: 'video/mp4', trimStartMs: 12000 });
    const pendingPlay = c.play().catch(() => {});
    const failure = await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('timeout')), 5000); c.subscribe(s => { if (s.status === 'error') { clearTimeout(timer); resolve(s.failure); } }); });
    await pendingPlay; const finalFailure = c.state.failure; c.dispose(); return { errors, failure, finalFailure, children: host.children.length };
  }, movie);
  assert.deepEqual(result, { errors: ['TypeError', 'TypeError', 'TypeError'], failure: 'invalid-duration', finalFailure: 'invalid-duration', children: 0 });
});

test('looping stays paused after a seek, repeats only while playing, and subscriptions are isolated', { timeout: 15000 }, async t => {
  const { page, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const options = { kind: 'video', source: new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' }), contentType: 'video/mp4', trimStartMs: 2000, trimEndMs: 2000, loop: true, muted: true };
    const c = adapter.mountBrowserMedia(host, options); options.loop = false;
    await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('timeout')), 5000); c.subscribe(s => { if (s.status === 'ready') { clearTimeout(timer); resolve(); } }); });
    let mediaPaused = true; Object.defineProperty(c.element, 'paused', { get: () => mediaPaused });
    c.element.play = async () => { mediaPaused = false; c.element.dispatchEvent(new Event('playing')); };
    c.element.pause = () => { mediaPaused = true; c.element.dispatchEvent(new Event('pause')); };
    let calls = 0; const off = c.subscribe(() => { calls++; }); c.subscribe(() => { throw Error('observer'); });
    c.seek(10000); c.element.dispatchEvent(new Event('timeupdate'));
    const paused = { time: c.element.currentTime, status: c.state.status, paused: c.element.paused };
    await c.play(); c.seek(10000); c.element.dispatchEvent(new Event('timeupdate'));
    const playing = { time: c.element.currentTime, status: c.state.status, paused: c.element.paused };
    off(); const after = calls; c.pause(); c.element.dispatchEvent(new Event('timeupdate')); const unsubscribed = calls === after;
    c.dispose(); return { paused, playing, unsubscribed };
  }, movie);
  assert.deepEqual(result.paused, { time: 2, status: 'paused', paused: true });
  assert.deepEqual(result.playing, { time: 2, status: 'playing', paused: false }); assert.equal(result.unsubscribed, true);
});

test('real Chromium playback loops a short trimmed interval without restarting a paused instance', { timeout: 12000 }, async t => {
  const { page, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const c = adapter.mountBrowserMedia(host, { kind: 'video', source: new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' }), contentType: 'video/mp4', trimStartMs: 2000, trimEndMs: 9400, loop: true, muted: true });
    try {
      await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('metadata timeout')), 5000); c.subscribe(s => { if (s.status === 'ready') { clearTimeout(timer); resolve(); } else if (s.status === 'error') { clearTimeout(timer); reject(Error(s.failure)); } }); });
      let max = 0, repeated = false;
      c.element.addEventListener('timeupdate', () => { const time = c.element.currentTime; if (max > 2.3 && time < max - 0.2) repeated = true; max = Math.max(max, time); });
      await c.play();
      const deadline = performance.now() + 5000;
      while (!repeated && performance.now() < deadline) await new Promise(resolve => setTimeout(resolve, 50));
      c.pause(); await new Promise(resolve => setTimeout(resolve, 80)); const paused = c.element.paused;
      return { repeated, max, paused, status: c.state.status };
    } finally { c.dispose(); }
  }, movie);
  assert.equal(result.repeated, true); assert.ok(result.max >= 2.3); assert.equal(result.paused, true); assert.equal(result.status, 'paused');
});

test('audio blobs play independently and stale native events do not change a paused state', async t => {
  const { page, requests } = await setup(t);
  const archive = readGeneratedZip(await readFile('tests/fixtures/pptkit-mixed-av.pptx'));
  const sound = [...archive.entries()].find(([name]) => name.endsWith('.wav'))[1];
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const c = adapter.mountBrowserMedia(host, { kind: 'audio', source: new Blob([Uint8Array.from(bytes)], { type: 'audio/wav' }), contentType: 'audio/wav', muted: true });
    try {
      await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('metadata timeout')), 5000); c.subscribe(s => { if (s.status === 'ready') { clearTimeout(timer); resolve(); } else if (s.status === 'error') { clearTimeout(timer); reject(Error(s.failure)); } }); });
      await c.play(); c.pause(); c.element.dispatchEvent(new Event('pause')); c.element.dispatchEvent(new Event('playing'));
      return { tag: c.element.tagName, paused: c.element.paused, status: c.state.status, capability: c.state.capability };
    } finally { c.dispose(); }
  }, [...sound]);
  assert.equal(result.tag, 'AUDIO'); assert.equal(result.paused, true); assert.equal(result.status, 'paused'); assert.ok(['maybe', 'probably'].includes(result.capability)); assert.deepEqual(requests, []);
});

test('media entry imports in Node without touching a DOM', async () => {
  const { mountBrowserMedia } = await import('../packages/svg-renderer/dist/media.js');
  assert.equal(typeof mountBrowserMedia, 'function');
  assert.throws(() => mountBrowserMedia(undefined, {}), /connected browser container/);
});

test('invalid bytes report actual media failure and disposal while loading revokes the URL', async t => {
  const { page, requests } = await setup(t);
  const result = await page.evaluate(async () => {
    const host = document.createElement('div'); document.body.append(host);
    const originalCreate = URL.createObjectURL, originalRevoke = URL.revokeObjectURL, live = new Set();
    URL.createObjectURL = x => { const u = originalCreate.call(URL, x); live.add(u); return u; };
    URL.revokeObjectURL = u => { live.delete(u); originalRevoke.call(URL, u); };
    try {
      const source = new Blob(['not valid wave'], { type: 'audio/wav' });
      const failed = adapter.mountBrowserMedia(host, { kind: 'audio', source, contentType: 'audio/wav' });
      const failure = await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('media error timeout')), 5000); failed.subscribe(s => { if (s.status === 'error') { clearTimeout(timer); resolve(s.failure); } }); });
      const cancelled = adapter.mountBrowserMedia(host, { kind: 'audio', source, contentType: 'audio/wav' });
      const states = []; cancelled.subscribe(s => states.push(s.status)); cancelled.dispose();
      await new Promise(resolve => setTimeout(resolve, 100));
      const cancellation = { states, status: cancelled.state.status, live: live.size, children: host.children.length };
      failed.dispose(); return { failure, cancellation, live: live.size, children: host.children.length };
    } finally { URL.createObjectURL = originalCreate; URL.revokeObjectURL = originalRevoke; }
  });
  assert.ok(['unsupported', 'decode'].includes(result.failure));
  assert.deepEqual(result.cancellation, { states: ['loading', 'disposed'], status: 'disposed', live: 1, children: 1 });
  assert.equal(result.live, 0); assert.equal(result.children, 0); assert.deepEqual(requests, []);
});

test('disposal is safe when media callbacks reenter disposal', async t => {
  const { page } = await setup(t);
  const result = await page.evaluate(() => {
    const host = document.createElement('div'); document.body.append(host);
    const c = adapter.mountBrowserMedia(host, { kind: 'audio', source: new Blob(['data'], { type: 'audio/wav' }), contentType: 'audio/wav' });
    let pauses = 0, disposedEvents = 0;
    c.element.pause = () => { pauses++; c.dispose(); };
    c.subscribe(s => { if (s.status === 'disposed') { disposedEvents++; c.dispose(); } });
    c.dispose(); return { pauses, disposedEvents, status: c.state.status, children: host.children.length };
  });
  assert.deepEqual(result, { pauses: 1, disposedEvents: 1, status: 'disposed', children: 0 });
});
