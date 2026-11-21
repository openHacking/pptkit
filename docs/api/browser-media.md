# Explicit browser media adapter

`@pptkit/svg-renderer/media` exports `mountBrowserMedia(container, options)`. This optional DOM entry mounts a single HTML audio or video element in a **connected** caller-owned container. It accepts an already acquired nonempty `Blob` and an explicit MIME/codec hint; it does not resolve asset IDs, read files, fetch URLs, start playback, create UI, or modify Core/Layout/standalone SVG output. `resolveAndMountBrowserMedia` additionally invokes an explicitly supplied resolver; only that caller-owned callback can load external resources. An application is responsible for authorization and loading before calling it. The element is exposed for styling, placement and optional native controls; the caller must not replace its source or remove it while the controller is active.

```ts
import { mountBrowserMedia } from '@pptkit/svg-renderer/media';

const media = mountBrowserMedia(container, {
  kind: 'video', source: videoBlob, contentType: 'video/mp4',
  muted: true, trimStartMs: 2000, trimEndMs: 500, controls: true,
});
const unsubscribe = media.subscribe(state => console.log(state.status, state.capability));
// In a user-initiated event handler:
await media.play();
// On leaving the page:
unsubscribe();
media.dispose();
```

`canPlayType` produces `probably`, `maybe` or `unsupported` as an advisory capability hint, independently of actual `play()` success. A rejected play promise is propagated and classified as `permission`, `unsupported` or `playback`; a permission/playback rejection may be retried with another `play()` call. Media-element errors separately report network, decode, unsupported or generic playback failure. Subscriptions receive an immediate frozen snapshot and subsequent state updates; observer exceptions do not affect other observers. Consumers should handle rejected `play()` promises and supply their own fallback and accessible controls.

Trim values are nonnegative source-relative milliseconds removed from the beginning/end. Once finite metadata is available, the interval must be nonempty. `seek` takes a source-relative time inside the retained interval and rejects before metadata is available. `volume` is a number from 0 to 1, default 0.8; mute, looping, return-to-start and native controls default to false. A user gesture remains the consumer's responsibility. Loop and end enforcement use browser `timeupdate`/`ended` events, so they are not sample-accurate and may overshoot, be delayed in background tabs or differ from PowerPoint. Exact timing, autoplay sequences, captions, implicit source loading, and slide navigation are not implemented by this adapter.

Each mount creates its own object URL and element even when two mounts share a Blob. `dispose()` is idempotent: it pauses, removes owned listeners and element, clears its source, revokes its URL, publishes `disposed`, and rejects subsequent controller operations. Consumers must dispose on leave, cancellation, failed loads and unmount. Disposing one instance does not revoke another instance's URL; no cross-slide continuity is implied. This does not cancel a caller's earlier Blob acquisition.

Browser tests exercise real Chromium video loading, playback, trim, independent mounts, cleanup, invalid duration, no network, plus controlled rejection and looping events. They do not establish codec support or playback behavior in other browsers, exact trim boundaries, subtitles or native PowerPoint equivalence.

## Resolver and lifetime cancellation

`resolveAndMountBrowserMedia(container, options)` accepts the same media settings except `source`, plus `resolve({ signal })` and a required positive integer `maxBytes`. It validates configuration before calling the resolver, snapshots the settings, and accepts only a nonempty Blob no larger than the limit. Resolver errors propagate; retry is a new explicit call. The optional `onResolveState` callback receives frozen `loading`, `mounted`, `cancelled`, or `failed` snapshots; callback exceptions are isolated. A pre-aborted signal reports only `cancelled`; invalid configuration rejects before reporting. `mounted` means the element was attached, not that media metadata loaded or playback succeeded. After mounting, observe `controller.subscribe` for playback state. If a `mounted` observer aborts synchronously, the returned controller is already disposed. The signal passed to the resolver is frozen into its context. An abort rejects promptly even if the resolver ignores it, and a late resolution cannot mount an element or create a URL. An already-aborted signal prevents the resolver from running.

```ts
const lifetime = new AbortController();
const media = await resolveAndMountBrowserMedia(container, {
  kind: 'audio', contentType: 'audio/wav', maxBytes: 4_000_000,
  signal: lifetime.signal,
  resolve: ({ signal }) => loadAuthorizedMedia(signal), // application-owned policy
});
// On page exit, or to cancel a pending resolve:
lifetime.abort();
```

A supplied signal must belong to the container's browser window. After mounting, aborting it disposes that instance and publishes `disposed`; explicit `dispose()` detaches the abort listener. Distinct mounts keep separate URLs and lifetimes. The size check is made **after** the resolver returns a Blob: it is not a streaming/download budget. The resolver must enforce authorization, URL/redirect policy, decompression and actual transfer limits itself. The adapter never fetches an address on its own, cannot stop a resolver that ignores its signal, and does not share or cache fetched blobs. Chromium tests cover an explicit fetch resolver, failure and retry, settings mutation, size rejection, cancellation with a late result, abort during/after mounting, and twenty repeated mount/leave/cancel cycles with owned listeners and object URLs counted.

## Composing with the bounded HTTP resolver

Applications that already use `@pptkit/pptx-exporter` can explicitly reuse its [HTTP media resolver](byte-assets.md#external-audio-and-video) without introducing an exporter dependency into the SVG renderer. Use the **same byte limit** for the stream and the returned Blob, and validate the response MIME according to the consumer's accepted formats:

```ts
import { createHttpMediaResolver } from '@pptkit/pptx-exporter';
import { resolveAndMountBrowserMedia } from '@pptkit/svg-renderer/media';

const limit = 4_000_000;
const lifetime = new AbortController();
const permittedUrl = 'https://media.example.test/sound.wav';
const http = createHttpMediaResolver({ authorizeUrl: url => url === permittedUrl });
const controller = await resolveAndMountBrowserMedia(container, {
  kind: 'audio', contentType: 'audio/wav', maxBytes: limit, signal: lifetime.signal,
  onResolveState: state => reportLoading(state.status),
  async resolve({ signal }) {
    const result = await http({
      id: 'sound', kind: 'audio', source: { type: 'url', value: permittedUrl },
      accessibility: { decorative: false },
    }, { signal, maxBytes: limit });
    if (!result || (result.mimeType && result.mimeType !== 'audio/wav')) throw new TypeError('Unexpected media type');
    return new Blob([new Uint8Array(result.bytes)], { type: result.mimeType ?? 'audio/wav' });
  },
});
```

The HTTP resolver counts actual response bytes and asks the supplied policy before requests. Redirects default to zero; browsers can expose manual redirects as opaque, in which case they are rejected. Authorization and URL/DNS policy remain application responsibilities. A local-server Chromium test covers authorized playback, denied requests, streaming oversize, opaque redirect rejection and cancellation of an incomplete response. No HTTP request is made unless the caller invokes the HTTP resolver.
