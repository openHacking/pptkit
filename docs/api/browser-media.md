# Explicit browser media adapter

`@pptkit/svg-renderer/media` exports `mountBrowserMedia(container, options)`. This optional DOM entry mounts a single HTML audio or video element in a **connected** caller-owned container. It accepts an already acquired nonempty `Blob` and an explicit MIME/codec hint; it does not resolve asset IDs, read files, fetch URLs, start playback, create UI, or modify Core/Layout/standalone SVG output. An application is responsible for authorization and loading before calling it. The element is exposed for styling, placement and optional native controls; the caller must not replace its source or remove it while the controller is active.

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

Trim values are nonnegative source-relative milliseconds removed from the beginning/end. Once finite metadata is available, the interval must be nonempty. `seek` takes a source-relative time inside the retained interval and rejects before metadata is available. `volume` is a number from 0 to 1, default 0.8; mute, looping, return-to-start and native controls default to false. A user gesture remains the consumer's responsibility. Loop and end enforcement use browser `timeupdate`/`ended` events, so they are not sample-accurate and may overshoot, be delayed in background tabs or differ from PowerPoint. Exact timing, autoplay sequences, captions, source loading, and slide navigation are not implemented by this adapter.

Each mount creates its own object URL and element even when two mounts share a Blob. `dispose()` is idempotent: it pauses, removes owned listeners and element, clears its source, revokes its URL, publishes `disposed`, and rejects subsequent controller operations. Consumers must dispose on leave, cancellation, failed loads and unmount. Disposing one instance does not revoke another instance's URL; no cross-slide continuity is implied. This does not cancel a caller's earlier Blob acquisition.

Browser tests exercise real Chromium video loading, playback, trim, independent mounts, cleanup, invalid duration, no network, plus controlled rejection and looping events. They do not establish codec support or playback behavior in other browsers, exact trim boundaries, subtitles or native PowerPoint equivalence.
