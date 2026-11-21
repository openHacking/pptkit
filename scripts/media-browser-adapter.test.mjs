import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createServer } from 'node:http';
import { once } from 'node:events';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { createPresentation } from '../packages/core/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
import { renderLayoutToSvg } from '../packages/svg-renderer/dist/index.js';

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

test('native media failure wins over a late play rejection and stale metadata', async t => {
  const { page, requests, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const c = adapter.mountBrowserMedia(host, { kind: 'video', source: new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' }), contentType: 'video/mp4' });
    await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('metadata timeout')), 5000); c.subscribe(s => { if (s.status === 'ready') { clearTimeout(timer); resolve(); } }); });
    let rejectPlay; c.element.play = () => new Promise((_, reject) => { rejectPlay = reject; });
    const pending = c.play().catch(e => e.name);
    Object.defineProperty(c.element, 'error', { configurable: true, value: { code: 3 } }); c.element.dispatchEvent(new Event('error'));
    rejectPlay(new DOMException('late rejection', 'NotAllowedError')); const rejected = await pending;
    c.element.dispatchEvent(new Event('loadedmetadata')); c.element.dispatchEvent(new Event('pause'));
    let retry; try { await c.play(); } catch (e) { retry = e.name; }
    const state = { status: c.state.status, failure: c.state.failure }; c.dispose();
    return { rejected, retry, state, remaining: host.children.length };
  }, movie);
  assert.deepEqual(result, { rejected: 'NotAllowedError', retry: 'TypeError', state: { status: 'error', failure: 'decode' }, remaining: 0 }); assert.deepEqual(requests, []);
});

test('permission failure survives metadata and retries stay loading until metadata is available', async t => {
  const { page, requests, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const c = adapter.mountBrowserMedia(host, { kind: 'video', source: new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' }), contentType: 'video/mp4' });
    c.element.dispatchEvent(new Event('pause')); const initial = c.state.status;
    c.element.play = () => Promise.reject(new DOMException('gesture', 'NotAllowedError'));
    await c.play().catch(() => {}); const rejected = c.state.failure;
    let resolvePlay; c.element.play = () => new Promise(resolve => { resolvePlay = resolve; });
    const pending = c.play(); const retrying = c.state.status; resolvePlay(); await pending;
    c.element.play = () => Promise.reject(new DOMException('gesture', 'NotAllowedError')); await c.play().catch(() => {});
    Object.defineProperty(c.element, 'duration', { configurable: true, value: 12 }); c.element.dispatchEvent(new Event('loadedmetadata'));
    const afterMetadata = { status: c.state.status, failure: c.state.failure, duration: c.state.durationMs }; c.dispose();
    return { initial, rejected, retrying, afterMetadata };
  }, movie);
  assert.deepEqual(result, { initial: 'loading', rejected: 'permission', retrying: 'loading', afterMetadata: { status: 'error', failure: 'permission', duration: 12000 } }); assert.deepEqual(requests, []);
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

test('resolver cancellation wins over noncooperative late completion without mounting', async t => {
  const { page, requests, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const originalCreate = URL.createObjectURL, originalRevoke = URL.revokeObjectURL, urls = new Set();
    URL.createObjectURL = value => { const url = originalCreate.call(URL, value); urls.add(url); return url; };
    URL.revokeObjectURL = url => { urls.delete(url); originalRevoke.call(URL, url); };
    try {
      const ac = new AbortController(); let finish, calls = 0, seenSignal;
      const promise = adapter.resolveAndMountBrowserMedia(host, { kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, signal: ac.signal,
        resolve({ signal }) { calls++; seenSignal = signal === ac.signal; return new Promise(resolve => { finish = resolve; }); } });
      await Promise.resolve(); ac.abort();
      let error; try { await promise; } catch (e) { error = e.name; }
      finish(new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' })); await Promise.resolve(); await Promise.resolve();
      return { error, calls, seenSignal, children: host.children.length, urls: urls.size };
    } finally { URL.createObjectURL = originalCreate; URL.revokeObjectURL = originalRevoke; }
  }, movie);
  assert.deepEqual(result, { error: 'AbortError', calls: 1, seenSignal: true, children: 0, urls: 0 }); assert.deepEqual(requests, []);
});

test('resolver validates before side effects, retries after failure, and uses a snapshot of settings', async t => {
  const { page, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const source = new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' }); let calls = 0;
    const resolver = () => { calls++; return source; };
    const errors = [];
    for (const options of [
      { kind: 'video', contentType: 'video/mp4', maxBytes: 0, resolve: resolver },
      { kind: 'video', contentType: 'video/mp4', maxBytes: 1, volume: 3, resolve: resolver },
      { kind: 'video', contentType: 'video/mp4', maxBytes: 1, resolve: resolver },
      { kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, resolve: () => { calls++; throw Error('failed'); } },
    ]) { try { await adapter.resolveAndMountBrowserMedia(host, options); } catch (e) { errors.push(e.name); } }
    let release;
    const options = { kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, muted: true, volume: 0.3, resolve: () => new Promise(resolve => { release = resolve; }) };
    const pending = adapter.resolveAndMountBrowserMedia(host, options); await Promise.resolve();
    options.muted = false; options.volume = 0.9; options.maxBytes = 1; release(source);
    const controller = await pending; const values = { muted: controller.element.muted, volume: controller.element.volume, children: host.children.length };
    controller.dispose(); return { errors, calls, values, remaining: host.children.length };
  }, movie);
  assert.deepEqual(result, { errors: ['TypeError', 'TypeError', 'RangeError', 'Error'], calls: 2, values: { muted: true, volume: 0.3, children: 1 }, remaining: 0 });
});

test('lifetime signal removes only its own mounted instance and handles already-aborted signals', async t => {
  const { page, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const source = new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' });
    const firstSignal = new AbortController(), secondSignal = new AbortController(); let calls = 0;
    const options = signal => ({ kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, muted: true, signal, resolve: () => { calls++; return source; } });
    const first = await adapter.resolveAndMountBrowserMedia(host, options(firstSignal.signal));
    const second = await adapter.resolveAndMountBrowserMedia(host, options(secondSignal.signal));
    const seen = []; first.subscribe(s => seen.push(s.status)); firstSignal.abort();
    const after = { first: first.state.status, second: second.state.status, children: host.children.length, seen: seen.at(-1) };
    let preaborted; try { await adapter.resolveAndMountBrowserMedia(host, options(firstSignal.signal)); } catch (e) { preaborted = e.name; }
    second.dispose(); secondSignal.abort();
    return { after, preaborted, calls, children: host.children.length };
  }, movie);
  assert.equal(result.after.first, 'disposed'); assert.notEqual(result.after.second, 'disposed'); assert.equal(result.after.children, 1); assert.equal(result.after.seen, 'disposed');
  assert.equal(result.preaborted, 'AbortError'); assert.equal(result.calls, 2); assert.equal(result.children, 0);
});

test('an explicitly supplied resolver can fetch media and the adapter issues no additional requests', async t => {
  const { page, requests } = await setup(t);
  const archive = readGeneratedZip(await readFile('tests/fixtures/pptkit-mixed-av.pptx'));
  const sound = [...archive.entries()].find(([name]) => name.endsWith('.wav'))[1];
  const allowed = [];
  await page.route('https://media.invalid/sound.wav', async route => { allowed.push(route.request().url()); await route.fulfill({ status: 200, contentType: 'audio/wav', headers: { 'access-control-allow-origin': '*' }, body: Buffer.from(sound) }); });
  const result = await page.evaluate(async () => {
    const host = document.createElement('div'); document.body.append(host); const ac = new AbortController();
    const controller = await adapter.resolveAndMountBrowserMedia(host, { kind: 'audio', contentType: 'audio/wav', maxBytes: 1000000, signal: ac.signal, muted: true,
      async resolve({ signal }) { const response = await fetch('https://media.invalid/sound.wav', { signal }); if (!response.ok) throw Error('fetch failed'); return response.blob(); } });
    const ready = await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('timeout')), 5000); controller.subscribe(state => { if (state.status === 'ready') { clearTimeout(timer); resolve(true); } else if (state.status === 'error') { clearTimeout(timer); reject(Error(state.failure)); } }); });
    await controller.play(); ac.abort(); return { ready, status: controller.state.status, children: host.children.length };
  });
  assert.deepEqual(result, { ready: true, status: 'disposed', children: 0 });
  assert.deepEqual(allowed, ['https://media.invalid/sound.wav']); assert.deepEqual(requests, []);
});

test('aborting while mounting rolls back and an already aborted signal never calls the resolver', async t => {
  const { page, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host); const signal = new AbortController();
    const original = host.appendChild.bind(host); host.appendChild = node => { const value = original(node); signal.abort(); return value; };
    const source = new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' }); let first, second, calls = 0;
    try { adapter.mountBrowserMedia(host, { kind: 'video', source, contentType: 'video/mp4', signal: signal.signal }); } catch (e) { first = e.name; }
    try { await adapter.resolveAndMountBrowserMedia(host, { kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, signal: signal.signal, resolve() { calls++; return source; } }); } catch (e) { second = e.name; }
    return { first, second, calls, children: host.children.length };
  }, movie);
  assert.deepEqual(result, { first: 'AbortError', second: 'AbortError', calls: 0, children: 0 });
});

test('immediate abort skips resolver and listener ownership ends after disposal', async t => {
  const { page, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host); const source = new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' });
    const first = new AbortController(); let calls = 0;
    const cancelled = adapter.resolveAndMountBrowserMedia(host, { kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, signal: first.signal, resolve: () => { calls++; return source; } });
    first.abort(); let aborted; try { await cancelled; } catch (e) { aborted = e.name; } await Promise.resolve();
    const second = new AbortController(), nativeAdd = second.signal.addEventListener.bind(second.signal), nativeRemove = second.signal.removeEventListener.bind(second.signal), active = new Set();
    second.signal.addEventListener = (name, listener, options) => { if (name === 'abort') active.add(listener); return nativeAdd(name, listener, options); };
    second.signal.removeEventListener = (name, listener, options) => { if (name === 'abort') active.delete(listener); return nativeRemove(name, listener, options); };
    const c = await adapter.resolveAndMountBrowserMedia(host, { kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, signal: second.signal, resolve: () => source });
    const afterMount = active.size; c.dispose(); const afterDispose = active.size;
    return { calls, aborted, afterMount, afterDispose, children: host.children.length };
  }, movie);
  assert.deepEqual(result, { calls: 0, aborted: 'AbortError', afterMount: 1, afterDispose: 0, children: 0 });
});

test('repeated mount, leave and cancelled loads leave no media, owned listeners or URLs', async t => {
  const { page, requests, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const source = new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' });
    const add = HTMLMediaElement.prototype.addEventListener, remove = HTMLMediaElement.prototype.removeEventListener;
    const create = URL.createObjectURL, revoke = URL.revokeObjectURL, liveUrls = new Set(), owned = new Map(); let played = 0;
    URL.createObjectURL = value => { const u = create.call(URL, value); liveUrls.add(u); return u; };
    URL.revokeObjectURL = url => { liveUrls.delete(url); revoke.call(URL, url); };
    HTMLMediaElement.prototype.addEventListener = function (name, fn, options) { const set = owned.get(this) ?? new Set(); set.add(fn); owned.set(this, set); return add.call(this, name, fn, options); };
    HTMLMediaElement.prototype.removeEventListener = function (name, fn, options) { owned.get(this)?.delete(fn); return remove.call(this, name, fn, options); };
    try {
      for (let i = 0; i < 20; i++) {
        const lifetime = new AbortController();
        const controller = await adapter.resolveAndMountBrowserMedia(host, { kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, signal: lifetime.signal, resolve: () => source });
        if (!controller.element.paused || controller.state.status === 'playing') played++;
        lifetime.abort();
        let late; const loading = new AbortController();
        const pending = adapter.resolveAndMountBrowserMedia(host, { kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, signal: loading.signal, resolve: () => new Promise(resolve => { late = resolve; }) });
        await Promise.resolve(); loading.abort(); await pending.catch(() => {}); late(source);
      }
      await Promise.resolve(); await Promise.resolve();
      return { children: host.children.length, liveUrls: liveUrls.size, listeners: [...owned.values()].reduce((sum, set) => sum + set.size, 0), played };
    } finally { HTMLMediaElement.prototype.addEventListener = add; HTMLMediaElement.prototype.removeEventListener = remove; URL.createObjectURL = create; URL.revokeObjectURL = revoke; }
  }, movie);
  assert.deepEqual(result, { children: 0, liveUrls: 0, listeners: 0, played: 0 }); assert.deepEqual(requests, []);
});

test('resolution events distinguish loading, failure, cancellation and mounted playback state', async t => {
  const { page, movie } = await setup(t);
  const result = await page.evaluate(async bytes => {
    const host = document.createElement('div'); document.body.append(host);
    const source = new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' });
    const make = (notify, resolve, signal) => ({ kind: 'video', contentType: 'video/mp4', maxBytes: 1000000, resolve, onResolveState: notify, ...(signal ? { signal } : {}) });
    const successful = [], controller = await adapter.resolveAndMountBrowserMedia(host, make(state => { successful.push([state.status, Object.isFrozen(state)]); throw Error('observer'); }, () => source));
    const playback = controller.state.status; controller.dispose();
    const failure = []; let failureName;
    try { await adapter.resolveAndMountBrowserMedia(host, make(s => failure.push(s.status), () => { throw new Error('resolver failed'); })); } catch (e) { failureName = e.message; }
    const cancellation = [], ac = new AbortController(); let called = 0, cancellationName;
    try { await adapter.resolveAndMountBrowserMedia(host, make(s => { cancellation.push(s.status); if (s.status === 'loading') ac.abort(); }, () => { called++; return source; }, ac.signal)); } catch (e) { cancellationName = e.name; }
    const pre = []; try { await adapter.resolveAndMountBrowserMedia(host, make(s => pre.push(s.status), () => source, ac.signal)); } catch {}
    const reentrant = [], ended = new AbortController();
    const disposed = await adapter.resolveAndMountBrowserMedia(host, make(s => { reentrant.push(s.status); if (s.status === 'mounted') ended.abort(); }, () => source, ended.signal));
    return { successful, playback, failure, failureName, cancellation, cancellationName, called, pre, reentrant, disposed: disposed.state.status, children: host.children.length };
  }, movie);
  assert.deepEqual(result.successful, [['loading', true], ['mounted', true]]); assert.equal(result.playback, 'loading');
  assert.deepEqual(result.failure, ['loading', 'failed']); assert.equal(result.failureName, 'resolver failed');
  assert.deepEqual(result.cancellation, ['loading', 'cancelled']); assert.equal(result.cancellationName, 'AbortError'); assert.equal(result.called, 0); assert.deepEqual(result.pre, ['cancelled']);
  assert.deepEqual(result.reentrant, ['loading', 'mounted']); assert.equal(result.disposed, 'disposed'); assert.equal(result.children, 0);
});

test('authorized streaming HTTP resolver composes with playback, budgets and cancellation', { timeout: 15000 }, async t => {
  const archive = readGeneratedZip(await readFile('tests/fixtures/pptkit-mixed-av.pptx'));
  const wave = Buffer.from([...archive.entries()].find(([name]) => name.endsWith('.wav'))[1]);
  const hits = []; let slowClosed = false;
  const server = createServer((request, response) => {
    hits.push({ path: request.url, cookie: request.headers.cookie ?? null, referrer: request.headers.referer ?? null });
    if (request.url === '/') { response.writeHead(200, { 'content-type': 'text/html' }); response.end('<!doctype html>'); return; }
    if (request.url === '/redirect') { response.writeHead(302, { location: '/sound' }); response.end(); return; }
    if (request.url === '/slow') { response.writeHead(200, { 'content-type': 'audio/wav' }); response.write(wave.subarray(0, 100)); response.on('close', () => { slowClosed = true; }); return; }
    if (request.url === '/sound') { response.writeHead(200, { 'content-type': 'audio/wav' }); response.write(wave.subarray(0, 100)); response.end(wave.subarray(100)); return; }
    response.writeHead(404); response.end();
  });
  server.listen(0, '127.0.0.1'); await once(server, 'listening'); t.after(() => { server.closeAllConnections(); server.close(); });
  const base = `http://127.0.0.1:${server.address().port}`;
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(); await page.goto(base);
  for (const name of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const source = await readFile('packages/svg-renderer/dist/browser/media.js', 'utf8');
  const result = await page.evaluate(async ({ source, base, limit }) => {
    const moduleUrl = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
    const { resolveAndMountBrowserMedia } = await import(moduleUrl); URL.revokeObjectURL(moduleUrl);
    const host = document.createElement('div'); document.body.append(host); const authorized = [];
    const resolver = PPTKit.createHttpMediaResolver({ maxRedirects: 1, authorizeUrl(url) { authorized.push(url); return !url.endsWith('/denied'); } });
    const mount = (path, maxBytes, signal, events) => resolveAndMountBrowserMedia(host, {
      kind: 'audio', contentType: 'audio/wav', maxBytes, ...(signal ? { signal } : {}), muted: true,
      onResolveState: state => events.push(state.status),
      async resolve({ signal }) {
        const asset = { id: 'sound', kind: 'audio', source: { type: 'url', value: base + path }, accessibility: { decorative: false } };
        const result = await resolver(asset, { signal, maxBytes });
        if (!result || result.mimeType && result.mimeType !== 'audio/wav') throw new TypeError('Unexpected media type');
        return new Blob([result.bytes], { type: result.mimeType ?? 'audio/wav' });
      },
    });
    const okEvents = [], controller = await mount('/sound', limit, undefined, okEvents);
    await new Promise((resolve, reject) => { const timeout = setTimeout(() => reject(Error('metadata timeout')), 5000); controller.subscribe(s => { if (s.status === 'ready') { clearTimeout(timeout); resolve(); } else if (s.status === 'error') { clearTimeout(timeout); reject(Error(s.failure)); } }); });
    await controller.play(); const played = !controller.element.paused; controller.dispose();
    const errors = [], statuses = [];
    for (const [path, max] of [['/denied', limit], ['/sound', limit - 1], ['/redirect', limit]]) {
      const events = []; try { await mount(path, max, undefined, events); } catch (e) { errors.push(e.message); } statuses.push(events);
    }
    const ac = new AbortController(), cancelledEvents = []; let cancelled;
    const pending = mount('/slow', limit, ac.signal, cancelledEvents); await new Promise(resolve => setTimeout(resolve, 150)); ac.abort();
    try { await pending; } catch (e) { cancelled = e.name; }
    return { played, okEvents, errors, statuses, cancelled, cancelledEvents, authorized, children: host.children.length };
  }, { source, base, limit: wave.length });
  assert.equal(result.played, true); assert.deepEqual(result.okEvents, ['loading', 'mounted']);
  assert.match(result.errors[0], /not authorized/); assert.match(result.errors[1], /maxBytes/); assert.match(result.errors[2], /not an accessible 200/);
  assert.deepEqual(result.statuses, [['loading', 'failed'], ['loading', 'failed'], ['loading', 'failed']]);
  assert.equal(result.cancelled, 'AbortError'); assert.deepEqual(result.cancelledEvents, ['loading', 'cancelled']); assert.equal(result.children, 0);
  assert.deepEqual(result.authorized, [base + '/sound', base + '/denied', base + '/sound', base + '/redirect', base + '/slow']);
  assert.deepEqual(hits.map(x => x.path), ['/', '/sound', '/sound', '/redirect', '/slow']);
  assert.ok(hits.slice(1).every(x => x.cookie === null && x.referrer === null));
  for (let i = 0; i < 30 && !slowClosed; i++) await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(slowClosed, true);
});

async function mediaLayout() {
  const parts = readGeneratedZip(await readFile('tests/fixtures/pptkit-mixed-av.pptx'));
  const p = createPresentation();
  p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: parts.get('ppt/media/movie.mp4') } });
  p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: parts.get('ppt/media/sound.wav') } });
  p.addSlide({ id: 'page', elements: [
    { type: 'video', id: 'video-a', videoAssetId: 'movie', playbackStart: 'automatic', playbackVolume: 'low', muted: true, trimStartMs: 150, trimEndMs: 50, box: { x: 0, y: 0, width: 40, height: 40 } },
    { type: 'video', id: 'video-b', videoAssetId: 'movie', playbackStart: 'object-click', box: { x: 50, y: 0, width: 40, height: 40 } },
    { type: 'audio', id: 'sound-a', audioAssetId: 'sound', playbackVolume: 'mute', loop: true, box: { x: 100, y: 0, width: 40, height: 40 } },
  ] });
  return resolveLayout(p);
}

test('layout bridge mounts explicit media identities with independent authored settings', async t => {
  const { page, requests } = await setup(t), layout = await mediaLayout();
  const result = await page.evaluate(async input => {
    const host = document.createElement('div'); document.body.append(host); const snapshot = JSON.stringify(input);
    const options = { slideId: 'page', elementId: 'video-a', maxBytes: 1000000 };
    const pending = adapter.mountLayoutMedia(host, input, options); options.slideId = 'changed'; options.elementId = 'changed';
    const first = await pending;
    const second = await adapter.mountLayoutMedia(host, input, { slideId: 'page', elementId: 'video-b', maxBytes: 1000000 });
    const audio = await adapter.mountLayoutMedia(host, input, { slideId: 'page', elementId: 'sound-a', maxBytes: 1000000 });
    const ready = controller => new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('metadata timeout')), 5000); controller.subscribe(s => { if (s.status === 'ready') { clearTimeout(timer); resolve(); } else if (s.status === 'error') { clearTimeout(timer); reject(Error(s.failure)); } }); });
    await Promise.all([ready(first.controller), ready(second.controller), ready(audio.controller)]);
    const before = { ids: [first.slideId, first.elementId, first.assetId, first.kind, first.playbackStart, second.playbackStart], first: [first.controller.element.volume, first.controller.element.muted, first.controller.element.currentTime], second: [second.controller.element.volume, second.controller.element.muted, second.controller.element.paused], audio: [audio.controller.element.volume, audio.controller.element.muted], untouched: JSON.stringify(input) === snapshot, children: host.children.length };
    await first.controller.play(); const independent = !first.controller.element.paused && second.controller.element.paused && audio.controller.element.paused;
    first.controller.dispose(); const afterOne = host.children.length; second.controller.dispose(); audio.controller.dispose();
    return { before, independent, afterOne, remaining: host.children.length };
  }, layout);
  assert.deepEqual(result.before.ids, ['page', 'video-a', 'movie', 'video', 'automatic', 'object-click']);
  assert.deepEqual(result.before.first, [0.2, true, 0.15]); assert.deepEqual(result.before.second, [0.8, false, true]); assert.deepEqual(result.before.audio, [0.8, true]);
  assert.equal(result.before.untouched, true); assert.equal(result.before.children, 3); assert.equal(result.independent, true); assert.equal(result.afterOne, 2); assert.equal(result.remaining, 0); assert.deepEqual(requests, []);
});

test('layout bridge rejects invalid identity, hidden media, invalid sources and oversize before external loading', async t => {
  const { page, requests } = await setup(t), layout = await mediaLayout();
  const result = await page.evaluate(async input => {
    const host = document.createElement('div'); document.body.append(host); let calls = 0;
    const attempt = async (change, options = {}) => {
      const candidate = structuredClone(input); change(candidate); try { const result = await adapter.mountLayoutMedia(host, candidate, { slideId: 'page', elementId: 'video-a', maxBytes: 1000000, resolveExternal() { calls++; throw Error('unexpected resolver'); }, ...options }); result.controller.dispose(); return 'ok'; } catch (error) { return error.name; }
    };
    const errors = [];
    errors.push(await attempt(() => {}, { slideId: 'missing' }));
    errors.push(await attempt(l => { l.slides[0].elements.push(structuredClone(l.slides[0].elements[0])); }));
    errors.push(await attempt(l => { l.slides[0].elements[0].hidden = true; }));
    errors.push(await attempt(l => { l.assets[0].kind = 'audio'; }));
    errors.push(await attempt(l => { l.assets[0].source = { type: 'bytes', value: 'AB==', encoding: 'base64' }; }));
    errors.push(await attempt(l => { l.assets[0].source = { type: 'bytes', value: 'AAAA', encoding: 'bad' }; }));
    errors.push(await attempt(() => {}, { maxBytes: 10 }));
    errors.push(await attempt(l => { l.slides[0].elements[0].muted = null; }));
    errors.push(await attempt(l => { l.slides[0].elements[0].playbackVolume = null; }));
    errors.push(await attempt(() => {}, { contentType: 'video/webm' }));
    errors.push(await attempt(() => {}, { contentType: 'video/' }));
    errors.push(await attempt(l => { l.assets.push(structuredClone(l.assets[0])); }));
    errors.push(await attempt(l => { l.slides.push(structuredClone(l.slides[0])); }));
    errors.push(await attempt(l => { l.assets[0].source = { type: 'other', value: 'bad' }; }));
    errors.push(await attempt(l => { const cycle = { type: 'group', id: 'cycle', hidden: false, children: [] }; cycle.children.push(cycle); l.slides[0].elements.push(cycle); }));
    errors.push(await attempt(l => { l.slides[0].elements = Array(100001).fill(l.slides[0].elements[0]); }));
    const hiddenSlide = await attempt(l => { l.slides[0].hidden = true; });
    const nested = await attempt(l => { const video = l.slides[0].elements[0]; l.slides[0].elements[0] = { type: 'group', id: 'hidden-group', hidden: true, children: [video] }; });
    const external = structuredClone(input); external.assets[0].source = { type: 'url', value: 'https://example.invalid/movie.mp4' };
    let absent; try { await adapter.mountLayoutMedia(host, external, { slideId: 'page', elementId: 'video-a', maxBytes: 1000000 }); } catch (error) { absent = error.name; }
    return { errors, hiddenSlide, nested, absent, calls, children: host.children.length };
  }, layout);
  assert.deepEqual(result.errors, ['TypeError', 'TypeError', 'TypeError', 'TypeError', 'TypeError', 'TypeError', 'RangeError', 'TypeError', 'TypeError', 'TypeError', 'TypeError', 'TypeError', 'TypeError', 'TypeError', 'TypeError', 'RangeError']);
  assert.equal(result.hiddenSlide, 'ok'); assert.equal(result.nested, 'TypeError'); assert.equal(result.absent, 'TypeError'); assert.equal(result.calls, 0); assert.equal(result.children, 0); assert.deepEqual(requests, []);
});

test('layout bridge snapshots external source and settings, enforces budget and propagates cancellation', async t => {
  const { page, requests, movie } = await setup(t), layout = await mediaLayout();
  const result = await page.evaluate(async ({ input, bytes }) => {
    input.assets[0].source = { type: 'url', value: 'https://example.invalid/movie.mp4' };
    const host = document.createElement('div'); document.body.append(host); const lifetime = new AbortController(); let release, observed;
    const options = { slideId: 'page', elementId: 'video-a', maxBytes: 1000000, signal: lifetime.signal, resolveExternal(asset, context) { observed = [asset.id, asset.source.value, context.maxBytes, context.signal === lifetime.signal, Object.isFrozen(asset), Object.isFrozen(asset.source), Object.isFrozen(context)]; return new Promise(resolve => { release = resolve; }); } };
    const pending = adapter.mountLayoutMedia(host, input, options); await Promise.resolve();
    input.assets[0].source.value = 'https://changed.invalid/'; input.slides[0].elements[0].playbackVolume = 'high'; options.elementId = 'changed';
    release(new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' })); const mounted = await pending;
    const values = { observed, elementId: mounted.elementId, volume: mounted.controller.element.volume, muted: mounted.controller.element.muted };
    lifetime.abort(); let oversize; try { await adapter.mountLayoutMedia(host, input, { slideId: 'page', elementId: 'video-a', maxBytes: 10, resolveExternal: () => new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' }) }); } catch (e) { oversize = e.name; }
    return { values, status: mounted.controller.state.status, oversize, children: host.children.length };
  }, { input: layout, bytes: movie });
  assert.deepEqual(result.values, { observed: ['movie', 'https://example.invalid/movie.mp4', 1000000, true, true, true, true], elementId: 'video-a', volume: 0.2, muted: true });
  assert.equal(result.status, 'disposed'); assert.equal(result.oversize, 'RangeError'); assert.equal(result.children, 0); assert.deepEqual(requests, []);
});

async function nestedMediaSlide() {
  const parts = readGeneratedZip(await readFile('tests/fixtures/pptkit-mixed-av.pptx'));
  const p = createPresentation({ size: { width: 1000.1, height: 562.6 } });
  p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: parts.get('ppt/media/movie.mp4') } });
  p.addSlide({ id: 'overlay-slide', elements: [{ type: 'group', id: 'outer', box: { x: 100, y: 70, width: 400, height: 240 }, coordinateSize: { width: 200, height: 120 }, transform: { rotation: 22, flipH: true }, opacity: 0.7, children: [
    { type: 'group', id: 'inner', box: { x: 30, y: 15, width: 120, height: 80 }, coordinateSize: { width: 60, height: 40 }, transform: { rotation: -15, flipV: true }, opacity: 0.8, children: [
      { type: 'video', id: 'video', videoAssetId: 'movie', box: { x: 5, y: 10, width: 40, height: 20 }, transform: { rotation: 35, flipH: true }, opacity: 0.9, muted: true },
    ] },
  ] }] });
  const layout = resolveLayout(p), rendered = await renderLayoutToSvg(layout);
  return { layout, svg: rendered.slides[0].svg };
}

test('SVG media overlay follows nested Layout transforms at multiple viewport sizes', async t => {
  const { page, requests } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout, svg: text }) => {
    const wrapper = document.createElement('div'); document.body.append(wrapper); wrapper.innerHTML = text;
    const svg = wrapper.querySelector('svg'); svg.style.width = '800px'; svg.style.height = '450px';
    const snapshot = JSON.stringify(layout), mounted = await adapter.mountLayoutMediaOverlay(svg, layout, { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000 });
    const original = svg.querySelector('rect[stroke-dasharray="6 4"]'), overlay = svg.querySelector('foreignObject');
    const matrix = element => { const m = element.getScreenCTM(); return [m.a, m.b, m.c, m.d, m.e, m.f]; };
    const bounds = element => { const b = element.getBBox(); return [b.x, b.y, b.width, b.height]; };
    const first = { original: matrix(original), overlay: matrix(overlay), originalBox: bounds(original), overlayBox: bounds(overlay) };
    svg.style.width = '500px'; svg.style.height = '300px'; const second = { original: matrix(original), overlay: matrix(overlay) };
    const root = svg.querySelector('[data-pptkit-media-overlay="video"]');
    const opacities = [...root.querySelectorAll('g')].map(x => x.getAttribute('opacity'));
    const styles = [mounted.controller.element.style.width, mounted.controller.element.style.height, mounted.controller.element.style.objectFit];
    const unchanged = JSON.stringify(layout) === snapshot, childrenBefore = svg.querySelectorAll('[data-pptkit-media-overlay]').length;
    mounted.controller.dispose(); return { first, second, opacities, styles, unchanged, childrenBefore, childrenAfter: svg.querySelectorAll('[data-pptkit-media-overlay]').length };
  }, input);
  for (const key of ['first', 'second']) for (let i = 0; i < 6; i++) assert.ok(Math.abs(result[key].original[i] - result[key].overlay[i]) < 1e-6, `${key} matrix ${i}`);
  assert.deepEqual(result.first.originalBox, result.first.overlayBox); assert.deepEqual(result.opacities, ['0.7', '0.8', '0.9']);
  assert.deepEqual(result.styles, ['100%', '100%', 'fill']); assert.equal(result.unchanged, true); assert.equal(result.childrenBefore, 1); assert.equal(result.childrenAfter, 0); assert.deepEqual(requests, []);
});

test('SVG overlay rejects mismatched slides and invalid geometry without leaving DOM or loading media', async t => {
  const { page, requests } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout: source, svg: text }) => {
    const wrapper = document.createElement('div'); document.body.append(wrapper); wrapper.innerHTML = text; const svg = wrapper.querySelector('svg');
    const errors = []; let loads = 0;
    const attempt = async (modify, attrs = {}) => {
      const layout = structuredClone(source); modify(layout); svg.setAttribute('data-pptkit-slide-id', attrs.id ?? 'overlay-slide'); svg.setAttribute('viewBox', attrs.viewBox ?? '0 0 1000.1 562.6');
      try { const result = await adapter.mountLayoutMediaOverlay(svg, layout, { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000, resolveExternal() { loads++; throw Error('unexpected'); } }); result.controller.dispose(); return 'ok'; } catch (e) { return e.name; }
    };
    errors.push(await attempt(() => {}, { id: 'wrong' })); errors.push(await attempt(() => {}, { viewBox: '0 0 1000 562.6' }));
    errors.push(await attempt(l => { l.slides[0].elements[0].coordinateSize.width = 0; }));
    errors.push(await attempt(l => { l.slides[0].elements[0].coordinateSize.width = 1e-200; l.slides[0].elements[0].children[0].coordinateSize.width = 1e-200; }));
    errors.push(await attempt(l => { l.slides[0].elements[0].box.x = Number.MAX_VALUE; l.slides[0].elements[0].box.width = Number.MAX_VALUE; }));
    errors.push(await attempt(l => { l.slides[0].elements[0].children[0].children[0].box.width = 0; }));
    errors.push(await attempt(l => { l.slides[0].elements[0].children[0].hidden = true; }));
    errors.push(await attempt(l => { let child = l.slides[0].elements[0].children[0].children[0]; const template = l.slides[0].elements[0]; for (let i = 0; i < 257; i++) child = { ...template, id: `deep-${i}`, children: [child] }; l.slides[0].elements = [child]; }));
    errors.push(await attempt(l => { l.assets = []; }));
    errors.push(await attempt(l => { l.assets[0].source = { type: 'url', value: 'https://example.invalid/v.mp4' }; }));
    return { errors, loads, overlays: svg.querySelectorAll('[data-pptkit-media-overlay]').length };
  }, input);
  assert.deepEqual(result.errors, ['TypeError', 'TypeError', 'TypeError', 'RangeError', 'RangeError', 'TypeError', 'TypeError', 'RangeError', 'TypeError', 'Error']);
  assert.equal(result.loads, 1); assert.equal(result.overlays, 0); assert.deepEqual(requests, []);
});

test('overlay cancels pending external loads and removes itself when the lifetime ends', async t => {
  const { page, movie, requests } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout, svg: text, bytes }) => {
    const wrapper = document.createElement('div'); document.body.append(wrapper); wrapper.innerHTML = text; const svg = wrapper.querySelector('svg');
    layout.assets[0].source = { type: 'url', value: 'https://example.invalid/v.mp4' };
    const ac = new AbortController(); let complete;
    const pending = adapter.mountLayoutMediaOverlay(svg, layout, { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000, signal: ac.signal, resolveExternal: () => new Promise(resolve => { complete = resolve; }) });
    await Promise.resolve(); const during = svg.querySelectorAll('[data-pptkit-media-overlay]').length; ac.abort(); let error; try { await pending; } catch (e) { error = e.name; }
    complete(new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' })); await Promise.resolve();
    const after = svg.querySelectorAll('[data-pptkit-media-overlay]').length;
    const secondSignal = new AbortController(); const mounted = await adapter.mountLayoutMediaOverlay(svg, layout, { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000, signal: secondSignal.signal, resolveExternal: () => new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' }) });
    const mountedCount = svg.querySelectorAll('[data-pptkit-media-overlay]').length; secondSignal.abort();
    return { during, error, after, mountedCount, finalCount: svg.querySelectorAll('[data-pptkit-media-overlay]').length, status: mounted.controller.state.status };
  }, { ...input, bytes: movie });
  assert.deepEqual(result, { during: 1, error: 'AbortError', after: 0, mountedCount: 1, finalCount: 0, status: 'disposed' }); assert.deepEqual(requests, []);
});

test('overlay rejects duplicate pending or mounted identities and permits retry after disposal', async t => {
  const { page, movie } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout, svg: text, bytes }) => {
    const wrapper = document.createElement('div'); document.body.append(wrapper); wrapper.innerHTML = text; const svg = wrapper.querySelector('svg');
    layout.assets[0].source = { type: 'url', value: 'https://example.invalid/v.mp4' }; let complete, calls = 0;
    const options = { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000, resolveExternal: () => { calls++; return new Promise(resolve => { complete = resolve; }); } };
    const first = adapter.mountLayoutMediaOverlay(svg, layout, options); await Promise.resolve();
    let pendingError, mountedError; try { await adapter.mountLayoutMediaOverlay(svg, layout, options); } catch (e) { pendingError = e.name; }
    complete(new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' })); const mounted = await first;
    try { await adapter.mountLayoutMediaOverlay(svg, layout, options); } catch (e) { mountedError = e.name; }
    mounted.controller.dispose();
    const retry = await adapter.mountLayoutMediaOverlay(svg, layout, { ...options, resolveExternal: () => new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' }) });
    const count = svg.querySelectorAll('[data-pptkit-media-overlay]').length; retry.controller.dispose();
    return { pendingError, mountedError, calls, count, remaining: svg.querySelectorAll('[data-pptkit-media-overlay]').length };
  }, { ...input, bytes: movie });
  assert.deepEqual(result, { pendingError: 'TypeError', mountedError: 'TypeError', calls: 1, count: 1, remaining: 0 });
});

test('SVG overlay displays an actual playing video over its static poster in Chromium', async t => {
  const { page, requests } = await setup(t), input = await nestedMediaSlide();
  const archive = readGeneratedZip(await readFile('tests/fixtures/powerpoint-caption-playback.pptx'));
  input.layout.assets[0].source.value = Buffer.from([...archive.entries()].find(([name]) => name.endsWith('.mp4'))[1]).toString('base64');
  const outer = input.layout.slides[0].elements[0], inner = outer.children[0], video = inner.children[0]; outer.opacity = 1; inner.opacity = 1; video.opacity = 1;
  input.svg = (await renderLayoutToSvg(input.layout)).slides[0].svg;
  const point = await page.evaluate(async ({ layout, svg: text }) => {
    const wrapper = document.createElement('div'); document.body.append(wrapper); wrapper.innerHTML = text; const svg = wrapper.querySelector('svg');
    svg.style.width = '800px'; svg.style.height = '450px';
    const mounted = await adapter.mountLayoutMediaOverlay(svg, layout, { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000 });
    window.activeMedia = mounted.controller;
    await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('metadata timeout')), 5000); mounted.controller.subscribe(s => { if (s.status === 'ready') { clearTimeout(timer); resolve(); } else if (s.status === 'error') { clearTimeout(timer); reject(Error(s.failure)); } }); });
    await mounted.controller.play();
    await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('frame timeout')), 5000); mounted.controller.element.requestVideoFrameCallback(() => { clearTimeout(timer); resolve(); }); });
    const fo = svg.querySelector('foreignObject'), box = fo.getBBox(), m = fo.getScreenCTM();
    const pt = new DOMPoint(box.x + box.width / 2, box.y + box.height / 2).matrixTransform(m);
    return { x: Math.round(pt.x), y: Math.round(pt.y) };
  }, input);
  const screenshot = await page.screenshot();
  const pixel = await page.evaluate(async ({ bytes, point }) => {
    const image = new Image(), url = URL.createObjectURL(new Blob([Uint8Array.from(bytes)], { type: 'image/png' }));
    try { image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height; const ctx = canvas.getContext('2d'); ctx.drawImage(image, 0, 0); return [...ctx.getImageData(point.x, point.y, 1, 1).data]; }
    finally { URL.revokeObjectURL(url); window.activeMedia.dispose(); }
  }, { bytes: [...screenshot], point });
  assert.ok(pixel[2] > 200 && pixel[0] < 40 && pixel[1] < 40, `expected blue video pixel, got ${pixel}`); assert.deepEqual(requests, []);
});

test('overlay leaves the static poster visible while loading and on media decode failure', async t => {
  const { page, requests } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout, svg: text }) => {
    const host = document.createElement('div'); document.body.append(host); host.innerHTML = text; const svg = host.querySelector('svg');
    layout.assets[0].source.value = btoa('not a valid video');
    const mounted = await adapter.mountLayoutMediaOverlay(svg, layout, { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000 });
    const loading = mounted.controller.element.style.visibility;
    const failure = await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('error timeout')), 5000); mounted.controller.subscribe(s => { if (s.status === 'error') { clearTimeout(timer); resolve(s.failure); } }); });
    const hidden = mounted.controller.element.style.visibility, poster = svg.querySelector('rect[stroke-dasharray="6 4"]') !== null;
    mounted.controller.dispose(); return { loading, failure, hidden, poster, overlays: svg.querySelectorAll('[data-pptkit-media-overlay]').length };
  }, input);
  assert.equal(result.loading, 'hidden'); assert.ok(['unsupported', 'decode'].includes(result.failure)); assert.equal(result.hidden, 'hidden'); assert.equal(result.poster, true); assert.equal(result.overlays, 0); assert.deepEqual(requests, []);
});

test('overlay rolls back its mounted controller when later setup fails', async t => {
  const { page, requests } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout, svg: text }) => {
    const wrapper = document.createElement('div'); document.body.append(wrapper); wrapper.innerHTML = text; const svg = wrapper.querySelector('svg');
    const create = URL.createObjectURL, revoke = URL.revokeObjectURL, live = new Set();
    URL.createObjectURL = value => { const url = create.call(URL, value); live.add(url); return url; };
    URL.revokeObjectURL = url => { live.delete(url); revoke.call(URL, url); };
    Object.defineProperty(HTMLMediaElement.prototype, 'style', { configurable: true, get() { throw Error('style unavailable'); } });
    try {
      let error; try { await adapter.mountLayoutMediaOverlay(svg, layout, { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000 }); } catch (e) { error = e.message; }
      return { error, urls: live.size, overlays: svg.querySelectorAll('[data-pptkit-media-overlay]').length };
    } finally { delete HTMLMediaElement.prototype.style; URL.createObjectURL = create; URL.revokeObjectURL = revoke; }
  }, input);
  assert.deepEqual(result, { error: 'style unavailable', urls: 0, overlays: 0 }); assert.deepEqual(requests, []);
});

test('in-place media follows nested SVG frames and restores the exact original DOM', async t => {
  const { page, requests } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout, svg: text }) => {
    const wrapper = document.createElement('div'); document.body.append(wrapper); wrapper.innerHTML = text; const svg = wrapper.querySelector('svg');
    svg.style.width = '800px'; svg.style.height = '450px'; const before = svg.outerHTML;
    const poster = [...svg.querySelectorAll('[data-pptkit-element-id]')].find(x => x.getAttribute('data-pptkit-element-id') === 'video');
    const mounted = await adapter.mountLayoutMediaInPlace(svg, layout, { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000 });
    const root = svg.querySelector('[data-pptkit-media-overlay="video"]'), fo = root.querySelector('foreignObject');
    const mat = element => { const m = element.getScreenCTM(); return [m.a, m.b, m.c, m.d, m.e, m.f]; };
    const matrix = [mat(poster), mat(fo)], adjacent = poster.nextElementSibling === root;
    await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('metadata timeout')), 5000); mounted.controller.subscribe(s => { if (s.status === 'ready') { clearTimeout(timer); resolve(); } else if (s.status === 'error') { clearTimeout(timer); reject(Error(s.failure)); } }); });
    const ready = { poster: poster.style.visibility, media: mounted.controller.element.style.visibility, label: mounted.controller.element.getAttribute('aria-label') };
    mounted.controller.dispose(); return { matrix, adjacent, ready, restored: svg.outerHTML === before, overlays: svg.querySelectorAll('[data-pptkit-media-overlay]').length };
  }, input);
  for (let i = 0; i < 6; i++) assert.ok(Math.abs(result.matrix[0][i] - result.matrix[1][i]) < 1e-6, `matrix ${i}`);
  assert.equal(result.adjacent, true); assert.equal(result.ready.poster, 'hidden'); assert.equal(result.ready.media, 'visible'); assert.ok(result.ready.label); assert.equal(result.restored, true); assert.equal(result.overlays, 0); assert.deepEqual(requests, []);
});

test('in-place media preserves drawing order and avoids translucent poster bleed', async t => {
  const { page, requests } = await setup(t), archive = readGeneratedZip(await readFile('tests/fixtures/powerpoint-caption-playback.pptx'));
  const p = createPresentation({ size: { width: 400, height: 250 } });
  p.registerAsset({ id: 'blue', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: [...archive.entries()].find(([name]) => name.endsWith('.mp4'))[1] } });
  p.addSlide({ id: 'layered', elements: [
    { type: 'video', id: 'blue-video', videoAssetId: 'blue', muted: true, opacity: 0.5, box: { x: 40, y: 40, width: 200, height: 100 } },
    { type: 'shape', id: 'red-shape', shape: 'rect', box: { x: 140, y: 40, width: 100, height: 100 }, style: { fill: { type: 'solid', color: 'FF0000' } } },
  ] });
  const layout = resolveLayout(p), svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const points = await page.evaluate(async ({ layout, svg: text }) => {
    const wrapper = document.createElement('div'); document.body.append(wrapper); wrapper.innerHTML = text; const svg = wrapper.querySelector('svg'); svg.style.width = '800px'; svg.style.height = '500px';
    const mounted = await adapter.mountLayoutMediaInPlace(svg, layout, { slideId: 'layered', elementId: 'blue-video', maxBytes: 1000000 }); window.activeMedia = mounted.controller;
    await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('metadata timeout')), 5000); mounted.controller.subscribe(s => { if (s.status === 'ready') { clearTimeout(timer); resolve(); } else if (s.status === 'error') { clearTimeout(timer); reject(Error(s.failure)); } }); });
    await mounted.controller.play(); await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('frame timeout')), 5000); mounted.controller.element.requestVideoFrameCallback(() => { clearTimeout(timer); resolve(); }); });
    const matrix = svg.getScreenCTM(), transform = (x, y) => { const pt = new DOMPoint(x, y).matrixTransform(matrix); return [Math.round(pt.x), Math.round(pt.y)]; };
    return [transform(90, 90), transform(190, 90)];
  }, { layout, svg });
  const screenshot = await page.screenshot();
  const pixels = await page.evaluate(async ({ bytes, points }) => {
    const image = new Image(), url = URL.createObjectURL(new Blob([Uint8Array.from(bytes)], { type: 'image/png' }));
    try { image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height; const ctx = canvas.getContext('2d'); ctx.drawImage(image, 0, 0); return points.map(([x, y]) => [...ctx.getImageData(x, y, 1, 1).data]); }
    finally { URL.revokeObjectURL(url); window.activeMedia.dispose(); }
  }, { bytes: [...screenshot], points });
  assert.ok(pixels[0][0] >= 120 && pixels[0][0] <= 135 && pixels[0][1] >= 120 && pixels[0][1] <= 135 && pixels[0][2] > 245, `unexpected transparent blue ${pixels[0]}`);
  assert.ok(pixels[1][0] > 245 && pixels[1][1] < 10 && pixels[1][2] < 10, `upper artwork must remain red: ${pixels[1]}`); assert.deepEqual(requests, []);
});

test('in-place media rejects stale frames and cross-mode duplicates without loading', async t => {
  const { page, requests } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout, svg: text }) => {
    const host = document.createElement('div'); document.body.append(host); host.innerHTML = text; const svg = host.querySelector('svg');
    const poster = [...svg.querySelectorAll('[data-pptkit-element-id]')].find(n => n.getAttribute('data-pptkit-element-id') === 'video');
    const options = { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000 }; const errors = [];
    const attempt = async (value = layout) => { try { const mounted = await adapter.mountLayoutMediaInPlace(svg, value, options); mounted.controller.dispose(); return 'ok'; } catch (e) { return e.name; } };
    const stale = structuredClone(layout); stale.slides[0].elements[0].children[0].children[0].box.width++; errors.push(await attempt(stale));
    const action = structuredClone(layout); action.slides[0].elements[0].children[0].children[0].action = { type: 'url', url: 'https://example.invalid/' }; errors.push(await attempt(action));
    poster.setAttribute('style', 'opacity: .5'); errors.push(await attempt()); poster.removeAttribute('style');
    const duplicate = poster.cloneNode(true); poster.parentNode.appendChild(duplicate); errors.push(await attempt()); duplicate.remove();
    const inPlace = await adapter.mountLayoutMediaInPlace(svg, layout, options);
    try { await adapter.mountLayoutMediaOverlay(svg, layout, options); errors.push('ok'); } catch (e) { errors.push(e.name); } inPlace.controller.dispose();
    const overlay = await adapter.mountLayoutMediaOverlay(svg, layout, options); errors.push(await attempt()); overlay.controller.dispose();
    return { errors, overlays: svg.querySelectorAll('[data-pptkit-media-overlay]').length };
  }, input);
  assert.deepEqual(result, { errors: Array(6).fill('TypeError'), overlays: 0 }); assert.deepEqual(requests, []);
});

test('in-place media restores its poster on decode failure and cancelled resolution', async t => {
  const { page, movie, requests } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout, svg: text, bytes }) => {
    const host = document.createElement('div'); document.body.append(host); host.innerHTML = text; const svg = host.querySelector('svg'); const before = svg.outerHTML;
    const options = { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000 };
    layout.assets[0].source.value = btoa('not valid video');
    const mounted = await adapter.mountLayoutMediaInPlace(svg, layout, options);
    const failure = await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('error timeout')), 5000); mounted.controller.subscribe(s => { if (s.status === 'error') { clearTimeout(timer); resolve(s.failure); } }); });
    const poster = [...svg.querySelectorAll('[data-pptkit-element-id]')].find(n => n.getAttribute('data-pptkit-element-id') === 'video');
    const fallback = !poster.hasAttribute('style') && mounted.controller.element.style.visibility === 'hidden'; mounted.controller.dispose();
    layout.assets[0].source = { type: 'url', value: 'https://example.invalid/v.mp4' }; const signal = new AbortController(); let complete;
    const pending = adapter.mountLayoutMediaInPlace(svg, layout, { ...options, signal: signal.signal, resolveExternal: () => new Promise(resolve => { complete = resolve; }) });
    await Promise.resolve(); signal.abort(); let aborted; try { await pending; } catch (e) { aborted = e.name; }
    complete(new Blob([Uint8Array.from(bytes)], { type: 'video/mp4' })); await Promise.resolve();
    return { failure, fallback, aborted, restored: svg.outerHTML === before };
  }, { ...input, bytes: movie });
  assert.ok(['unsupported', 'decode'].includes(result.failure)); assert.equal(result.fallback, true); assert.equal(result.aborted, 'AbortError'); assert.equal(result.restored, true); assert.deepEqual(requests, []);
});

test('in-place poster remains restored when a pending play rejects after a decode error', async t => {
  const { page, requests } = await setup(t), input = await nestedMediaSlide();
  const result = await page.evaluate(async ({ layout, svg: text }) => {
    const host = document.createElement('div'); document.body.append(host); host.innerHTML = text; const svg = host.querySelector('svg'), before = svg.outerHTML;
    const mounted = await adapter.mountLayoutMediaInPlace(svg, layout, { slideId: 'overlay-slide', elementId: 'video', maxBytes: 1000000 });
    const c = mounted.controller, poster = [...svg.querySelectorAll('[data-pptkit-element-id]')].find(n => n.getAttribute('data-pptkit-element-id') === 'video');
    await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('metadata timeout')), 5000); c.subscribe(s => { if (s.status === 'ready') { clearTimeout(timer); resolve(); } else if (s.status === 'error') { clearTimeout(timer); reject(Error(s.failure)); } }); });
    let rejectPlay; c.element.play = () => new Promise((_, reject) => { rejectPlay = reject; }); const pending = c.play().catch(() => {});
    Object.defineProperty(c.element, 'error', { configurable: true, value: { code: 3 } }); c.element.dispatchEvent(new Event('error'));
    rejectPlay(new DOMException('late', 'NotAllowedError')); await pending; c.element.dispatchEvent(new Event('loadedmetadata'));
    const after = { status: c.state.status, failure: c.state.failure, poster: getComputedStyle(poster).visibility, media: c.element.style.visibility }; c.dispose();
    return { after, restored: svg.outerHTML === before };
  }, input);
  assert.deepEqual(result, { after: { status: 'error', failure: 'decode', poster: 'visible', media: 'hidden' }, restored: true }); assert.deepEqual(requests, []);
});
