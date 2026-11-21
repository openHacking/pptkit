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

A native media error invalidates any pending play request for state reporting: its later rejection still reaches that request's caller but cannot replace the newer native failure. A recoverable permission or playback failure remains observable if metadata arrives afterward; a retry before metadata returns to `loading`, so an SVG poster is not hidden prematurely. Decode, network, unsupported and invalid-duration failures require a fresh mount.

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

A supplied signal must belong to the container's browser window. After mounting, aborting it disposes that instance and publishes `disposed`; explicit `dispose()` detaches the abort listener. Distinct mounts keep separate URLs and lifetimes. The size check is made **after** the resolver returns a Blob: it is not a streaming/download budget. The resolver must enforce authorization, URL/redirect policy, decompression and actual transfer limits itself. The adapter never directly fetches an address, cannot stop a resolver that ignores its signal, and does not share or cache fetched blobs. Browser media engines may themselves follow references inside formats they support; use an appropriate network sandbox or CSP when loading untrusted media. Chromium tests cover an explicit fetch resolver, failure and retry, settings mutation, size rejection, cancellation with a late result, abort during/after mounting, and twenty repeated mount/leave/cancel cycles with owned listeners and object URLs counted.

## Explicit shared loading

`createSharedBrowserMediaPool(window, { maxBytes, resolve })` optionally shares an in-flight or resolved Blob among pending and active mounts with the same caller-supplied key. Each mount still creates its own media element, controller and object URL, with independent playback, cancellation and disposal. The pool retains a Blob only while at least one mount using its key is pending or active. When the last user cancels, fails or disposes, the entry is removed and its resolver signal is aborted; a later mount resolves again. A resolver that ignores cancellation can still finish its own work, but its result cannot revive the removed entry. A failed or oversized result is not kept after the affected mounts settle.

```ts
import { createSharedBrowserMediaPool } from '@pptkit/svg-renderer/media';

const pool = createSharedBrowserMediaPool(window, {
  maxBytes: 4_000_000,
  resolve: (key, { signal, maxBytes }) => loadAuthorizedBlob(key, { signal, maxBytes }),
});
const first = await pool.mount(firstHost, 'movie-v1', { kind: 'video', contentType: 'video/mp4' });
const second = await pool.mount(secondHost, 'movie-v1', { kind: 'video', contentType: 'video/mp4' });
// Dispose each controller independently when its view leaves.
```

The pool belongs to one browser window. Choose a key that uniquely represents the authorized content and version: the pool does not inspect source identity, verify MIME, enforce authorization or separate incompatible uses sharing a key. The loader must enforce actual transfer and decompression limits; `maxBytes` checks the returned Blob and is also passed to the loader. The pool does not itself fetch, share object URLs, move a playing element between slides, or keep playback running after an instance is disposed. Chromium tests cover concurrent mounts, independent playback, one-user and last-user cancellation, late results, invalid/oversized data, separate keys, retry and synchronous abort on mounting.

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

## Mount a Layout media object

`mountLayoutMedia(container, layout, options)` connects a **resolved** `LayoutResult` to the same explicit browser controller. It looks up exactly one slide by `slideId`, exactly one slide-local audio/video object by `elementId`, and exactly one matching asset. It returns those identities, the media kind, the authored `playbackStart` intent and the controller. The intent is informative: even `automatic` does not start playback. The consumer owns slide selection, user gestures, geometry, transforms, layering and placement of the supplied container; this method does not overlay or mutate an SVG. Layout/master-inherited media are not searched. A hidden element or an element inside a hidden group is rejected; an explicitly selected hidden slide may be mounted. Group contents are searched without treating their local boxes as screen coordinates.

```ts
import { mountLayoutMedia } from '@pptkit/svg-renderer/media';

const mounted = await mountLayoutMedia(container, layout, {
  slideId: 'slide-1', elementId: 'video-1', maxBytes: 4_000_000,
  signal: lifetime.signal, onResolveState: state => reportLoading(state.status),
});
// Placement is application-owned; call mounted.controller.play() from a gesture.
```

Byte-backed Core assets are decoded from canonical Base64 with the caller's byte limit checked before decoding. A declared audio/video MIME is required unless the caller supplies `contentType`; a supplied MIME's base type must match an existing declaration. For path or URL sources there is no default loader: `resolveExternal(asset, { signal, maxBytes })` must be provided explicitly. The frozen request is detached from Layout, and the resolver owns authorization, stream/redirect budgets and MIME validation. The returned Blob is checked against `maxBytes`; it does not itself bound the resolver's transfer. Layout and option values are captured before asynchronous resolution. Audio `mute` maps to volume 0.8 with mute enabled; low, medium and high map to 0.2, 0.5 and 0.8. Video mute remains independent of its selected volume. Trim, loop and rewind use the same approximate browser semantics documented above.

Missing or ambiguous IDs/assets, wrong asset kind, malformed/oversize bytes, unsupported source types and invalid settings fail rather than silently choosing a target. Slide, asset and element traversal have fixed 100,000-item ceilings and reject cycles or repeated object references. The adapter uses independent URLs even when objects reference the same asset; dispose each controller or abort its lifetime. Browser tests cover mixed audio/video, independent instances, explicit external resolution, no network, identity conflicts, hidden objects, source mutation, limits and cancellation. They do not validate browser/native timing equivalence or geometry alignment.

## Explicit SVG overlay

`mountLayoutMediaOverlay(svg, layout, options)` is an optional convenience for a **connected SVG slide produced from the same Layout result**. It checks the slide ID and viewBox, then appends an SVG `foreignObject` with the selected media object's box, rotation, flips, group coordinate scales and group/object opacity. The same SVG viewBox transforms both the poster and the overlay as the slide is resized. Media still does not start automatically. A second pending or mounted overlay for the same object is rejected; dispose or abort the first before retrying. Loading failure, cancellation or controller disposal removes the adapter-owned overlay. The media element remains hidden while metadata loads and on a media error, leaving the static poster visible; it becomes visible when ready or playing.

```ts
import { mountLayoutMediaOverlay } from '@pptkit/svg-renderer/media';

const mounted = await mountLayoutMediaOverlay(svgSlide, layout, {
  slideId: 'slide-1', elementId: 'video-1', maxBytes: 4_000_000,
  signal: lifetime.signal, controls: true,
});
// In a user gesture, if appropriate: await mounted.controller.play();
```

The overlay is appended **above all existing SVG artwork**, not at the original drawing's z-order. The original static poster remains underneath: overlapping objects can be covered and translucent media can blend with that poster. Use this only where those differences are acceptable, or manage hiding/replacement and stacking in application code. The consumer must also provide accessible controls and labels and account for the SVG's `role="img"`, which can hide interactive descendants from assistive technology. This adapter does not alter the static renderer's accessibility tree or provide captions. Do not treat a live DOM with an overlay as a standalone static export.

The supplied SVG must retain the renderer's `data-pptkit-slide-id` and matching `viewBox`; the caller is responsible for ensuring its artwork came from the same Layout snapshot. Hidden media/hidden groups, invalid geometry, nonfinite composed transforms and group nesting beyond 256 levels are rejected. Chromium tests compare the overlay's screen transform and box with the static poster under nested rotations, flips, opacity and resizing; a captured playing frame verifies a blue video pixel at the transformed position. Tests also cover loading fallback, cancellation, duplicate mounts and cleanup. Browser-specific `foreignObject`, pointer and assistive-technology behavior, exact native stacking and translucent-poster replacement remain unverified.

## Explicit in-place SVG media

`mountLayoutMediaInPlace(svg, layout, options)` uses the same options and controller as the overlay above, but inserts the media beside its original frame in the SVG group so later artwork stays above it. The video poster is hidden only after the media is ready and restored on loading failure, playback error or disposal. Audio without native controls retains its poster. Native controls are opt-in; decorative media cannot expose them. The adapter supplies a label from the object's description or name, but the default SVG `role="img"` may hide interactive descendants from assistive technology. Applications must provide and test suitable accessible controls and semantics.

The SVG must be the connected, unmodified rendering of the same Layout snapshot. The adapter checks slide identity, viewBox, frame and group identities, geometry, transforms and opacity, and rejects unsupported actions, ambiguous frames, pre-existing inline styles and mismatched computed transforms. These checks cannot establish that all artwork or external styles match the snapshot; callers must maintain that invariant throughout the mount. Duplicate overlays, including the top-layer mode, are rejected. Dispose the controller or abort its lifetime when leaving the slide. The static renderer's output is unchanged, and no media starts automatically.

Chromium tests compare nested transforms and actual captured pixels for overlapping artwork and translucent video, and verify decode failure, cancellation, duplicate detection and exact DOM restoration. Browser-specific `foreignObject`, controls, pointer handling, assistive technology, captions and native PowerPoint playback fidelity remain unverified.
