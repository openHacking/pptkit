import { test } from 'node:test';
import assert from 'node:assert/strict';
import { captureImageDataUrl, ImageInspectionLimitError } from '../dist/rendering/image-data.js';

test('full data URL capture enforces canonical bytes and exact budget', () => {
  assert.deepEqual([...captureImageDataUrl('data:image/jpeg;base64,/9g=', 2)], [255, 216]);
  assert.deepEqual([...captureImageDataUrl('data:image/png,%00%ffA', 3)], [0, 255, 65]);
  assert.throws(() => captureImageDataUrl('data:image/png;base64,/9g=', 1), ImageInspectionLimitError);
  assert.throws(() => captureImageDataUrl('data:image/png,%00%ffA', 2), ImageInspectionLimitError);
  assert.deepEqual([...captureImageDataUrl('data:image/png,', 0)], []);
  for (const invalid of ['data:image/png;base64,/9h=', 'data:image/png;base64,/9g', 'data:image/png;base64,/9g= ', 'data:image/png,%zz', 'data:image/png,#x', 'data:image/png,é', 'https://example.invalid/x']) {
    assert.throws(() => captureImageDataUrl(invalid, 100), TypeError);
  }
});
