# JPEG EXIF orientation: named display comparison

A synthetic 300×200 JPEG has six colored cells and EXIF orientation 6. Pillow 12.1.1 independently decodes the JPEG and applies EXIF orientation to produce a 200×300 PNG control. PPTKit places each in the same 200×300 stretched frame and retains both resource bytes. Its opt-in JPEG diagnostic reports that orientation handling is unverified.

In the named macOS 26.5 (25F71) environment, **PowerPoint 16.104** opened the generated PPTX and displayed its first slide in slideshow mode. The JPEG displayed the unrotated 3-column × 2-row source pixel grid, stretched into its frame; the normalized PNG displayed the rotated 2-column × 3-row grid. In contrast, **Chromium 149.0.7827.55** displayed the JPEG and PNG preview as the same rotated grid, with network requests blocked. This is a concrete preview/native mismatch for this fixture, export and application version; it does not establish the behavior of other orientations, fit/crop modes, Office versions or producers.

The [independent oracle](../../scripts/research/oracles/jpeg-orientation-native.py) checks the JPEG orientation and dimensions, verifies the PNG equals Pillow's decoded EXIF-transposed pixels, follows the two PPTX image relationships and checks exact resource bytes and boxes, validates SVG data URLs, and classifies selected color-cell pixels in each retained screenshot. Three altered controls are rejected. [Evidence](jpeg-orientation-native-evidence.json) binds input, output, screenshots and source-script hashes; it does not independently witness the manual PowerPoint UI action or compare every pixel. The source, PPTX, SVG and producer report were reproduced byte-for-byte in a second directory. Local `.tmp` screenshots are not distributed.

To reproduce the generated source and Chromium rendering (Pillow and Playwright Chromium are research dependencies):

```sh
pnpm build
python3 scripts/prepare-jpeg-orientation-source.py .tmp/new-jpeg-orientation
node scripts/prepare-jpeg-orientation-native.mjs .tmp/new-jpeg-orientation
node scripts/inspect-jpeg-orientation-browser.mjs .tmp/new-jpeg-orientation
```

The named PowerPoint screenshot was collected manually through the computer UI. The oracle currently checks `.tmp/jpeg-orientation-native-01`: `python3 scripts/research/oracles/jpeg-orientation-native.py`. JPEG metadata reading and the warning do not normalize pixels; a consumer that needs matching output must supply an explicitly normalized image and account for its altered dimensions/crop policy. The default exporter does not automatically normalize; ICC/color fidelity and save/reopen behavior remain open under AUT-04.3.3.

## Explicit resolver normalization

The generated `normalized-candidate.pptx` uses the same authored frames and a resolver that explicitly returns the Pillow-normalized PNG for the JPEG asset with `orientationNormalization: "pixels"`. PowerPoint 16.104 opened it and slideshow displayed both sampled grids in the rotated 2×3 arrangement. The [normalization oracle](../../scripts/research/oracles/jpeg-orientation-normalized.py) verifies that the PNG equals the independently EXIF-transposed JPEG, both package image relationships contain that PNG, the two frames retain their expected geometry, and the screenshot's sampled colors match. It rejects three substituted source/package/screenshot controls. [Evidence](jpeg-orientation-normalized-evidence.json) binds the files and previous comparison; the UI action remains a manual observation. The source and normalized output bytes differ by design, and the exporter reports `asset-orientation-normalized`.

This proves the explicit caller-supplied conversion path for the named orientation-6 stretched example only. The exporter does not perform or verify the pixel conversion, and asset dimensions/crop are not automatically adjusted. Original metadata, ICC behavior, other EXIF orientations, native save/reopen and other fit modes remain unverified.


## Explicit browser decoder adapter

The optional `@pptkit/pptx-exporter/dom` adapter decodes JPEG bytes with the browser and encodes oriented PNG pixels. Chromium 149.0.7827.55 was tested against independently permuted raw pixel grids for all eight EXIF values, including mirrored orientations. For the named orientation-6 sample, its PNG matches Pillow's EXIF-transposed RGB pixels exactly. A PPTX produced through the resolver declaration displayed both the browser-normalized image and PNG control in the same rotated 2×3 arrangement in PowerPoint 16.104 slideshow.

The [oracle](../../scripts/research/oracles/jpeg-browser-normalized.py) checks image pixels, package relationships and geometry, sampled screenshot colors and three corrupted controls; [evidence](jpeg-browser-normalized-evidence.json) binds the artifacts and scripts. The native UI action is a manual observation; the screenshot is sampled, not a complete visual or save/reopen comparison. Firefox, Safari, other native applications, crop/fit behavior, ICC and metadata fidelity are not established. See the [adapter contract](../api/byte-assets.md#explicit-browser-normalization) for resource and cancellation limits.


## EXIF 6 mapped crop in PowerPoint

A raw JPEG crop excluding the rightmost third was mapped for orientation 6, giving a bottom crop of one third on the browser-normalized PNG. An independent Pillow control physically crops the corresponding raw 200×200 pixel region and rotates it clockwise. PowerPoint 16.104 opened the generated PPTX and slideshow displayed the mapped crop and physical control with the same yellow/red and magenta/green cells. This is a manual observation for one square crop and one application version; it does not establish other orientations, fit modes, native editing or save/reopen.

The [oracle](../../scripts/research/oracles/jpeg-crop-native.py) checks decoded and transformed source pixels, package image relationships and boxes, the serialized crop, screenshot cell samples and four rejected substitutions. [Evidence](jpeg-crop-native-evidence.json) binds source, screenshot, package and scripts; it cannot independently witness the UI action or establish complete pixel fidelity. The PNG control, PPTX and producer report reproduced byte-for-byte in a second directory.


A copy of that candidate was edited in PowerPoint by appending `R1` to its title, saved, closed and reopened. Its slideshow again showed the two matching sampled grids. PowerPoint omitted zero-valued crop attributes, but retained the bottom crop value, picture frames and exact embedded PNG bytes. PPTKit content import recovered the crop, dimensions and source bytes, and a no-op save was byte-identical to the native file. The [reopen oracle](../../scripts/research/oracles/jpeg-crop-reopen.py) and [evidence](jpeg-crop-reopen-evidence.json) bind the native file, screenshot, parser report and three rejected controls. The UI sequence remains a manual observation, and the screenshots are sampled. This is one title edit and one save/reopen, not an edit of the crop itself or a claim about other image geometries.


A further isolated copy was edited using PowerPoint's left crop handle, then saved, closed and reopened. PowerPoint changed the left source crop to 25.326% and narrowed/repositioned the picture frame; the independent control stayed unchanged. PPTKit content import recovered the edited crop, frame and exact embedded PNG bytes, and no-op save remained byte-identical. Chromium rendered the imported SVG preview with sampled color cells matching the reopened PowerPoint slideshow; the sampled horizontal image spans differed by one pixel. The [edit oracle](../../scripts/research/oracles/jpeg-crop-edit.py) checks the saved relationships, crop and geometry, both screenshots, parser/preview report and three negative controls; [evidence](jpeg-crop-edit-evidence.json) binds their hashes. UI provenance remains manual, and color and boundary samples are not a full visual comparison. This one left-handle edit does not establish all crop handles, rotations, other orientations or applications.
