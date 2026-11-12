# Named GIF playback observation

A repository-generated one-slide deck contains two picture objects sharing the same 40×40 red/blue GIF. Pillow 12.1.1 independently reads two frames, each with a stored 1,500 ms delay, disposal 2 and Netscape repeat count 0. The original source, generated PPTX, preview SVG and producer report were reproduced byte-for-byte in a second directory. The independent [oracle](../../scripts/research/oracles/gif-native-playback.py) checks the GIF colors and stored delays, two slide relationships to the same media part, exact packaged GIF bytes, preview data URLs, sample colors, ordering and five negative cases. [Evidence](gif-powerpoint-playback-evidence.json) records the source and screenshot hashes; local `.tmp` artifacts are not distributed.

On macOS 26.5 (25F71), **PowerPoint 16.104** opened the candidate and its slideshow was started via the UI. A first set of 18 screenshots captured a red-to-blue change. A second uninterrupted set of 60 screenshots captured **red → blue → red**, and both picture locations had the same sampled color in every screenshot. Separately, headless **Chromium 149.0.7827.55** displayed the SVG preview and 150 screenshots captured red → blue → red at both picture locations, with network requests blocked. This demonstrates visible change and a return to red for this sample in each named environment. It does not prove exact frame durations, continuous synchronization, indefinite looping, PowerPoint editing, repair-dialog absence in other workflows, reduced-motion behavior or other applications/versions. Native UI provenance is the recorded manual observation; the oracle inspects artifacts and cannot independently witness the UI actions.

To reproduce the source and automated browser sampling (Pillow and Playwright Chromium are research dependencies):

```sh
pnpm build
python3 scripts/prepare-gif-native-source.py .tmp/new-gif-playback
node scripts/prepare-gif-native-playback.mjs .tmp/new-gif-playback
node scripts/inspect-gif-browser-playback.mjs .tmp/new-gif-playback
```

The named PowerPoint capture used the original local directory `.tmp/run-_2s_4erk`. Its screenshots were collected through computer UI observation, not by the above browser script. The oracle currently checks that retained directory: `python3 scripts/research/oracles/gif-native-playback.py`. Sample screenshots contain solid color regions and no timing overlay; capture timestamps bracket each screenshot request, not display-frame transition times. Sampling can miss transitions, and the initial and second native batches have a gap.

## PowerPoint save and reopen

PowerPoint 16.104 performed three distinct Save As, close and exact-path reopen rounds on the named candidate. Independent ZIP/XML inspection of all three outputs confirms the same GIF bytes, both picture identities, descriptions, geometry and shared media target; PowerPoint renamed the media part and merged the two slide relationships into one. On the third reopened file, a new 90-image slideshow sample captured **blue → red → blue → red → blue**, with both sampled image locations agreeing in every capture. Earlier reopened files were not separately sampled for playback. The [reopen oracle](../../scripts/research/oracles/gif-native-reopen.py) validates the retained files and images and rejects four corrupt controls; its [evidence](gif-powerpoint-reopen-evidence.json) binds hashes. It checks selected package properties, not full package equivalence, and cannot witness the UI actions or measure exact timing.

The four files were imported through PPTKit's content projection in Node, browser ESM and browser global. Each projected one shared GIF byte asset and two image references; all twelve no-op saves matched their respective source bytes. The [parser evidence](gif-powerpoint-reopen-parser-evidence.json) records the scoped checks and source/build hashes. It does not establish editable GIF semantics or preservation for structural changes.

The export warning remains appropriate: one named playback result cannot establish general compatibility. Static fallback, other GIF constructs, zero-delay playback, finite loop counts, platform differences, editing and other native save/reopen samples remain open under AUT-04.3.2.
