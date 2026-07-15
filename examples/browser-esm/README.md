# Browser ESM Example

This zero-build example loads PPTKit from a CDN with an import map. Serve this
directory from any static HTTP server, open `index.html`, then preview the slides or
download the generated PPTX. PPTX generation happens entirely in the browser.

The checked example uses `@latest` so it follows the newest published browser bundle.
Production applications should replace every `@latest` with the same explicit PPTKit
version.

See the [browser CDN guide](../../docs/guides/browser-cdn.md) for the ESM and global
script contracts, security notes, and asset restrictions.
