# PowerPoint link activation attempt

On 2026-10-29, PowerPoint for Mac was used to inspect the two-page fixture in [the shared-layout package evidence](cross-layout-reuse-powerpoint-evidence.json). The original generated input and the third PowerPoint-saved file were opened separately. In each, the accessibility tree on slide 2 described `element-4` as containing a link to `Page A`; the independent package inspector found the slide relationship and `hlinkClick` in both files.

The slideshow started from `Page B`, and coordinate clicks aimed at its title were attempted. The original input and third saved file each advanced to the end-of-show screen rather than returning to `Page A`. For a further control, a copy of the original was modified only in slide 2 XML by replacing its `a:noFill` with an alpha-zero solid fill. Three attempts in that copy also reached the end-of-show screen. No successful positive control established that the coordinate input hit the intended object, so these observations do not prove a hyperlink defect or that transparent fill has no effect. Neither forward activation nor other apps or PowerPoint versions were tested.

The files in `.tmp/cross-layout-reuse-powerpoint/` had these SHA-256 hashes at the time of the attempt:

- `input.pptx`: `8868b51b9bf710499b38e8a3daa657b61a717b931b3c00a203cae505586171ea`
- `round3.pptx`: `51e18c5f7856e34bc33ebbd81a2084355a261dfd3f10e4d37c326f8aea4d5228`
- `transparent-control.pptx`: `043f2c6de656723a7941e26e6774da2c60cfb1c38f972986e9e252b094a207b4`

On a subsequent attempt in the same environment, keyboard navigation provided an independent activation path: from slide 2, `Tab` displayed a dotted focus rectangle around the linked title and `Return` moved the show back to slide 1. This was observed separately for the original input, the third saved file and the transparent-fill control. On the original, focusing the slide 1 title and pressing `Return` also displayed slide 2, although this forward result alone is indistinguishable from ordinary advancement. The backward result establishes keyboard activation for these named files, but does not establish a working mouse hit area. An additional coordinate click in the original title rectangle still ended the show; the input's mapping to screen coordinates has not been independently calibrated.

UI observations are recorded here; no persistent screenshot or automated replay was captured. A reliable positive mouse hit-test control and repeatable captures are needed before attributing the mouse result to the exporter or claiming mouse activation.
