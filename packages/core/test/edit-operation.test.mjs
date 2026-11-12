import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, PresentationEditError } from '../dist/index.js';

function fixture() {
  const doc = createPresentation({ id: 'deck' });
  doc.addSlide({ id: 'a' }); doc.addSlide({ id: 'b', hidden: true });
  return doc;
}
const op = changes => ({ schemaVersion: 'pptkit-edit-v1', documentId: 'deck', changes });
const hidden = (slideId, expected, value) => ({ type: 'slide-hidden', slideId, expected, value });
const number = (expected, value) => ({ type: 'first-slide-number', expected, value });
const order = (expected, value) => ({ type: 'slide-order', expected, value });
const textRun = (slideId, elementId, paragraphIndex, runIndex, expected, value) => ({ type: 'text-run', slideId, elementId, paragraphIndex, runIndex, expected, value });
const cellRun = (slideId, elementId, rowIndex, cellIndex, paragraphIndex, runIndex, expected, value) => ({ type: 'table-cell-text-run', slideId, elementId, rowIndex, cellIndex, paragraphIndex, runIndex, expected, value });
const state = d => [d.firstSlideNumber ?? null, ...d.slides.map(s => [s.id, s.hidden])];
const code = expected => error => error instanceof PresentationEditError && error.code === expected;

test('serializable batch applies atomically and inverse supports deterministic undo and redo', () => {
  const doc = fixture(), other = fixture(), original = state(doc);
  const request = op([hidden('a', false, true), number(null, 0), hidden('b', true, false)]);
  const result = doc.applyEditOperation(JSON.parse(JSON.stringify(request)));
  assert.deepEqual(state(doc), [0, ['a', true], ['b', false]]);
  assert.ok(Object.isFrozen(result) && Object.isFrozen(result.inverse.changes) && Object.isFrozen(result.changes[0]));
  assert.deepEqual(state(other), original);
  other.applyEditOperation(JSON.parse(JSON.stringify(request))); assert.deepEqual(state(other), state(doc));
  const redo = doc.applyEditOperation(JSON.parse(JSON.stringify(result.inverse))).inverse;
  assert.deepEqual(state(doc), original);
  doc.applyEditOperation(JSON.parse(JSON.stringify(redo))); assert.deepEqual(state(doc), state(other));
  assert.throws(() => doc.applyEditOperation(request), code('edit-conflict'));
  assert.deepEqual(state(doc), state(other));
});

test('conflicts, missing targets and wrong documents reject the complete batch', () => {
  const doc = fixture(), initial = state(doc);
  const cases = [
    [op([hidden('a', false, true), hidden('b', false, false)]), 'edit-conflict'],
    [op([hidden('a', false, true), number(3, 8)]), 'edit-conflict'],
    [op([hidden('a', false, true), hidden('missing', false, true)]), 'missing-target'],
    [{ ...op([hidden('a', false, true)]), documentId: 'elsewhere' }, 'wrong-document'],
  ];
  for (const [request, expected] of cases) { assert.throws(() => doc.applyEditOperation(request), code(expected)); assert.deepEqual(state(doc), initial); }
  doc.removeSlide('b'); assert.throws(() => doc.applyEditOperation(op([hidden('b', true, false)])), code('missing-target'));
});

test('operation capture rejects ambiguous data, duplicate targets and accessors before mutation', () => {
  const doc = fixture(), initial = state(doc); let calls = 0;
  const accessor = { type: 'slide-hidden', slideId: 'a', get expected() { calls++; return false; }, value: true };
  const sparse = new Array(1);
  const accessorArray = [hidden('a', false, true)]; Object.defineProperty(accessorArray, '0', { enumerable: true, get() { calls++; return hidden('a', false, true); } });
  const symbolChange = { ...hidden('a', false, true), [Symbol('extra')]: 1 };
  const cases = [op([accessor]), op([hidden('a', false, true), hidden('a', false, false)]), op([number(null, 1), number(null, 2)]),
    op([number(null, -1)]), op([number(null, NaN)]), op([number(null, undefined)]), op(sparse), op(accessorArray), op([symbolChange]),
    op([{ ...hidden('a', false, true), extra: true }]), { ...op([]), extra: 1 }, op([{ type: 'unknown' }]), op(new Array(10001))];
  for (const request of cases) { assert.throws(() => doc.applyEditOperation(request), code('invalid-operation')); assert.deepEqual(state(doc), initial); }
  assert.equal(calls, 0);
});

test('slide order participates in mixed atomic batches and inverse replay', () => {
  const doc = fixture(), initial = state(doc), a = doc.slides[0], b = doc.slides[1];
  const request = op([order(['a', 'b'], ['b', 'a']), hidden('a', false, true), number(null, 5)]);
  const result = doc.applyEditOperation(JSON.parse(JSON.stringify(request)));
  assert.deepEqual(state(doc), [5, ['b', true], ['a', true]]);
  assert.equal(doc.slides[0], b); assert.equal(doc.slides[1], a);
  assert.ok(Object.isFrozen(result.changes[0].expected) && Object.isFrozen(result.inverse.changes[0].value));
  const redo = doc.applyEditOperation(result.inverse).inverse;
  assert.deepEqual(state(doc), initial); doc.applyEditOperation(redo);
  assert.deepEqual(state(doc), [5, ['b', true], ['a', true]]);
  assert.throws(() => doc.applyEditOperation(request), code('edit-conflict'));
});

test('stale slide order and malformed permutations reject without partial changes', () => {
  const doc = fixture(), initial = state(doc); let invoked = 0;
  const badArray = ['a', 'b']; Object.defineProperty(badArray, '0', { enumerable: true, get() { invoked++; return 'a'; } });
  for (const change of [order(['b', 'a'], ['a', 'b']), order(['a'], ['a'])]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), change])), code('edit-conflict'));
    assert.deepEqual(state(doc), initial);
  }
  for (const changes of [[order(['a', 'b'], ['a', 'a'])], [order(['a', 'b'], ['a', 'c'])], [order(badArray, ['b', 'a'])], [order(['a', 'b'], ['b', 'a']), order(['a', 'b'], ['a', 'b'])]]) {
    assert.throws(() => doc.applyEditOperation(op(changes)), code('invalid-operation')); assert.deepEqual(state(doc), initial);
  }
  assert.equal(invoked, 0);
  doc.addSlide({ id: 'new' });
  assert.throws(() => doc.applyEditOperation(op([order(['a', 'b'], ['b', 'a'])])), code('edit-conflict'));
});

test('text run edits retain formatting and identity across slides, undo and redo', () => {
  const doc = fixture(), [a, b] = doc.slides;
  a.addElement({ id: 'one', type: 'text', content: [{ runs: [{ text: 'hello', style: { bold: true } }, { text: ' world' }] }] });
  b.addElement({ id: 'two', type: 'text', content: [{ runs: [{ text: 'other' }] }] });
  const before = structuredClone(a.elements[0]);
  const request = op([textRun('a', 'one', 0, 0, 'hello', 'hi'), textRun('a', 'one', 0, 1, ' world', '!'), textRun('b', 'two', 0, 0, 'other', 'else'), hidden('a', false, true), order(['a', 'b'], ['b', 'a'])]);
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(request))).inverse;
  assert.deepEqual(a.elements[0].content[0].runs.map(run => run.text), ['hi', '!']);
  assert.deepEqual(a.elements[0].content[0].runs[0].style, before.content[0].runs[0].style);
  assert.equal(a.elements[0].id, 'one'); assert.equal(b.elements[0].content[0].runs[0].text, 'else');
  const redo = doc.applyEditOperation(JSON.parse(JSON.stringify(undo))).inverse;
  assert.deepEqual(a.elements[0], before); assert.deepEqual(state(doc), [null, ['a', false], ['b', true]]);
  doc.applyEditOperation(redo); assert.deepEqual(a.elements[0].content[0].runs.map(run => run.text), ['hi', '!']);
});

test('text run failures preserve every edit and reject ambiguous targets', () => {
  const doc = fixture(), a = doc.slides[0];
  a.addElement({ id: 'one', type: 'text', content: [{ runs: [{ text: 'a' }, { text: 'b' }] }] });
  a.addElement({ id: 'plain', type: 'text', content: 'plain' });
  a.addElement({ id: 'malformed', type: 'text', content: [{}] });
  const before = structuredClone(a.elements), initial = state(doc);
  const cases = [
    [textRun('a', 'one', 0, 1, 'stale', 'x'), 'edit-conflict'],
    [textRun('a', 'one', 1, 0, '', 'x'), 'missing-target'],
    [textRun('a', 'plain', 0, 0, 'plain', 'x'), 'missing-target'],
    [textRun('a', 'malformed', 0, 0, '', 'x'), 'missing-target'],
    [textRun('missing', 'one', 0, 0, 'a', 'x'), 'missing-target'],
    [textRun('a', 'one', -1, 0, 'a', 'x'), 'invalid-operation'],
  ];
  for (const [bad, expected] of cases) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), textRun('a', 'one', 0, 0, 'a', 'changed'), bad])), code(expected));
    assert.deepEqual(a.elements, before); assert.deepEqual(state(doc), initial);
  }
  assert.throws(() => doc.applyEditOperation(op([textRun('a', 'one', 0, 0, 'a', 'x'), textRun('a', 'one', 0, 0, 'a', 'y')])), code('invalid-operation'));
  a.replaceElement('one', { ...a.elements[0], content: [{ runs: [{ text: 'still registered' }] }] });
});

test('text run edits preserve native field metadata and timing references', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'field', type: 'text', content: [{ runs: [{ text: '1', field: { id: 'field-id', kind: 'slide-number' } }] }] });
  slide.setEventTimeline({ stages: [{ id: 'stage', target: { pageId: 'a', elementId: 'field' }, trigger: { type: 'object-click', elementId: 'field' }, timing: { kind: 'span', id: 'time', durationMs: 10, fill: 'hold' } }] });
  const plan = slide.getEventTimelinePlan();
  const inverse = doc.applyEditOperation(op([textRun('a', 'field', 0, 0, '1', '2')])).inverse;
  assert.deepEqual(slide.elements[0].content[0].runs[0].field, { id: 'field-id', kind: 'slide-number' });
  assert.deepEqual(slide.getEventTimelinePlan(), plan);
  doc.applyEditOperation(inverse); assert.equal(slide.elements[0].content[0].runs[0].text, '1');
  assert.deepEqual(slide.getEventTimelinePlan(), plan);
});

test('nested shape and text edits in one group commit and undo together', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'group', type: 'group', coordinateSize: { width: 100, height: 100 }, box: { x: 0, y: 0, width: 100, height: 100 }, children: [
    { id: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 40, height: 40 }, text: { content: [{ runs: [{ text: 'shape', style: { italic: true } }] }] } },
    { id: 'inner', type: 'group', coordinateSize: { width: 40, height: 40 }, box: { x: 40, y: 0, width: 40, height: 40 }, children: [{ id: 'nested', type: 'text', content: [{ runs: [{ text: 'nested' }] }] }] },
  ] });
  slide.setEventTimeline({ stages: [{ id: 'stage', target: { pageId: 'a', elementId: 'nested' }, trigger: { type: 'object-click', elementId: 'shape' }, timing: { kind: 'span', id: 'time', durationMs: 10, fill: 'hold' } }] });
  const before = structuredClone(slide.elements[0]), plan = slide.getEventTimelinePlan();
  const changes = [textRun('a', 'shape', 0, 0, 'shape', 'new shape'), textRun('a', 'nested', 0, 0, 'nested', 'new nested')];
  assert.throws(() => doc.applyEditOperation(op([...changes, textRun('a', 'missing', 0, 0, '', 'x')])), code('missing-target'));
  assert.deepEqual(slide.elements[0], before);
  const inverse = doc.applyEditOperation(op(changes)).inverse, group = slide.elements[0];
  assert.equal(group.children[0].text.content[0].runs[0].text, 'new shape');
  assert.deepEqual(group.children[0].text.content[0].runs[0].style, { italic: true });
  assert.equal(group.children[1].children[0].content[0].runs[0].text, 'new nested');
  assert.deepEqual(slide.getEventTimelinePlan(), plan);
  doc.applyEditOperation(inverse); assert.deepEqual(slide.elements[0], before); assert.deepEqual(slide.getEventTimelinePlan(), plan);
  assert.throws(() => slide.addElement({ id: 'nested', type: 'text', content: 'duplicate' }), /Duplicate/);
});

test('table text edits use authored cell indexes, retain spans and styles, and undo atomically', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'group-table', type: 'group', coordinateSize: { width: 100, height: 100 }, box: { x: 0, y: 0, width: 100, height: 100 }, children: [{ id: 'table', type: 'table', box: { x: 0, y: 0, width: 100, height: 100 }, columns: [50, 50], rows: [
    { cells: [{ colSpan: 2, content: [{ runs: [{ text: 'merged', style: { bold: true } }] }] }] },
    { cells: [{ content: [{ runs: [{ text: 'left' }] }] }, { content: [{ runs: [{ text: 'right' }] }] }] },
  ] }] });
  const original = structuredClone(slide.elements[0]);
  const changes = [cellRun('a', 'table', 0, 0, 0, 0, 'merged', 'new'), cellRun('a', 'table', 1, 1, 0, 0, 'right', 'edited'), hidden('a', false, true)];
  const invalid = [cellRun('a', 'table', 0, 1, 0, 0, 'merged', 'x'), cellRun('a', 'table', 1, 0, 0, 0, 'stale', 'x'), cellRun('a', 'table', -1, 0, 0, 0, '', 'x')];
  for (const bad of invalid) {
    assert.throws(() => doc.applyEditOperation(op([...changes, bad])), code(bad.rowIndex === -1 ? 'invalid-operation' : bad.expected === 'stale' ? 'edit-conflict' : 'missing-target'));
    assert.deepEqual(slide.elements[0], original); assert.equal(slide.hidden, false);
  }
  assert.throws(() => doc.applyEditOperation(op([changes[0], changes[0]])), code('invalid-operation'));
  const inverse = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse;
  const table = slide.elements[0].children[0];
  assert.equal(table.rows[0].cells[0].colSpan, 2); assert.deepEqual(table.rows[0].cells[0].content[0].runs[0].style, { bold: true });
  assert.equal(table.rows[0].cells[0].content[0].runs[0].text, 'new'); assert.equal(table.rows[1].cells[1].content[0].runs[0].text, 'edited');
  const redo = doc.applyEditOperation(inverse).inverse; assert.deepEqual(slide.elements[0], original); assert.equal(slide.hidden, false);
  doc.applyEditOperation(redo); assert.equal(slide.elements[0].children[0].rows[1].cells[1].content[0].runs[0].text, 'edited');
});

test('string shorthand edits preserve representation across nested objects, cells and inverse replay', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'group', type: 'group', box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [
    { id: 'text', type: 'text', content: 'draft' },
    { id: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 50, height: 50 }, text: { content: 'shape', textStylePreset: 'preset' } },
    { id: 'table', type: 'table', box: { x: 0, y: 50, width: 100, height: 50 }, columns: [100], rows: [{ cells: [{ content: 'cell', colSpan: 1 }] }] },
  ] });
  const before = structuredClone(slide.elements[0]);
  const text = (elementId, expected, value) => ({ type: 'text-string', slideId: 'a', elementId, expected, value });
  const cell = (expected, value) => ({ type: 'table-cell-text-string', slideId: 'a', elementId: 'table', rowIndex: 0, cellIndex: 0, expected, value });
  const changes = [text('text', 'draft', 'final'), text('shape', 'shape', ''), cell('cell', 'new cell'), hidden('a', false, true)];
  for (const [bad, error] of [[cell('stale', 'x'), 'invalid-operation'], [text('missing', '', 'x'), 'missing-target'], [text('shape', 'stale', 'x'), 'invalid-operation'], [{ ...cell('cell', 'x'), rowIndex: -1 }, 'invalid-operation']]) {
    assert.throws(() => doc.applyEditOperation(op([...changes, bad])), code(error));
    assert.deepEqual(slide.elements[0], before); assert.equal(slide.hidden, false);
  }
  assert.throws(() => doc.applyEditOperation(op([text('text', 'wrong', 'x')])), code('edit-conflict'));
  assert.throws(() => doc.applyEditOperation(op([text('text', 'draft', 'final'), cell('stale', 'x')])), code('edit-conflict'));
  assert.throws(() => doc.applyEditOperation(op([textRun('a', 'text', 0, 0, 'draft', 'x')])), code('missing-target'));
  assert.throws(() => doc.applyEditOperation(op([text('text', 'draft', 'final'), textRun('a', 'text', 0, 0, 'draft', 'x')])), code('missing-target'));
  assert.deepEqual(slide.elements[0], before); assert.equal(slide.hidden, false);
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse;
  const group = slide.elements[0]; assert.equal(group.children[0].content, 'final'); assert.equal(group.children[1].text.content, '');
  assert.equal(group.children[1].text.textStylePreset, 'preset'); assert.equal(group.children[2].rows[0].cells[0].content, 'new cell');
  const redo = doc.applyEditOperation(undo).inverse; assert.deepEqual(slide.elements[0], before); assert.equal(slide.hidden, false);
  doc.applyEditOperation(redo); assert.equal(slide.elements[0].children[0].content, 'final');
});

test('element boxes compose with text edits, nested geometry, absent frames and auto height', () => {
  const doc = fixture(), slide = doc.slides[0];
  const groupBox = { x: 1, y: 2, width: 100, height: 100 }, oldShape = { x: 5, y: 6, width: 20, height: 30 }, newShape = { x: -5, y: 7, width: 0, height: 12 };
  slide.addElement({ id: 'group', type: 'group', box: groupBox, coordinateSize: { width: 100, height: 100 }, children: [
    { id: 'shape', type: 'shape', shape: 'rect', box: oldShape, text: { content: 'old' } },
    { id: 'auto', type: 'text', box: { x: 0, y: 0, width: 30 }, content: 'auto' },
    { id: 'absent', type: 'text', content: 'absent' },
  ] });
  const before = structuredClone(slide.elements[0]);
  const box = (id, expected, value) => ({ type: 'element-box', slideId: 'a', elementId: id, expected, value });
  const changes = [box('group', groupBox, { x: 10, y: 20, width: 200, height: 100 }), box('shape', oldShape, newShape),
    box('auto', { x: 0, y: 0, width: 30, height: null }, { x: 1, y: 2, width: 50, height: 40 }),
    box('absent', null, { x: 0, y: 0, width: 20, height: null }), { type: 'text-string', slideId: 'a', elementId: 'shape', expected: 'old', value: 'new' }, hidden('a', false, true)];
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse, current = slide.elements[0];
  assert.deepEqual(current.box, { x: 10, y: 20, width: 200, height: 100 }); assert.deepEqual(current.children[0].box, newShape);
  assert.equal(current.children[0].text.content, 'new'); assert.deepEqual(current.children[2].box, { x: 0, y: 0, width: 20 });
  const redo = doc.applyEditOperation(JSON.parse(JSON.stringify(undo))).inverse; assert.deepEqual(slide.elements[0], before); assert.equal(slide.hidden, false);
  doc.applyEditOperation(redo); assert.deepEqual(slide.elements[0].children[0].box, newShape);
});

test('box conflicts, malformed data, connector and unsupported auto height roll back the batch', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'shape', type: 'shape', shape: 'rect', box: { x: 1, y: 2, width: 3, height: 4 } });
  slide.addElement({ id: 'connector', type: 'connector', start: { x: 0, y: 0 }, end: { x: 2, y: 2 } });
  const before = structuredClone(slide.elements), valid = { x: 1, y: 2, width: 3, height: 4 };
  const box = (id, expected, value) => ({ type: 'element-box', slideId: 'a', elementId: id, expected, value });
  let invoked = 0; const getter = { x: 1, y: 2, width: 3, get height() { invoked++; return 4; } };
  for (const [bad, codeName] of [[box('shape', { ...valid, x: 9 }, valid), 'edit-conflict'], [box('connector', null, valid), 'missing-target'], [box('shape', valid, { ...valid, height: null }), 'invalid-operation'], [box('shape', valid, { ...valid, width: -1 }), 'invalid-operation'], [box('shape', valid, getter), 'invalid-operation'], [box('shape', valid, { ...valid, x: Infinity }), 'invalid-operation']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(codeName));
    assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  }
  assert.equal(invoked, 0); assert.throws(() => doc.applyEditOperation(op([box('shape', valid, valid), box('shape', valid, valid)])), code('invalid-operation'));
});

test('element visibility and opacity preserve absent versus authored defaults and compose with nested edits', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'group', type: 'group', box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [
    { id: 'text', type: 'text', box: { x: 0, y: 0, width: 50, height: 20 }, content: 'old', hidden: false, opacity: 1 },
  ] });
  const before = structuredClone(slide.elements[0]);
  const property = (type, id, expected, value) => ({ type, slideId: 'a', elementId: id, expected, value });
  const changes = [property('element-hidden', 'group', null, true), property('element-opacity', 'group', null, 0.5),
    property('element-hidden', 'text', false, null), property('element-opacity', 'text', 1, 0),
    { type: 'text-string', slideId: 'a', elementId: 'text', expected: 'old', value: 'new' }, hidden('a', false, true)];
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse, group = slide.elements[0];
  assert.equal(group.hidden, true); assert.equal(group.opacity, 0.5); assert.equal(group.children[0].opacity, 0);
  assert.equal(Object.hasOwn(group.children[0], 'hidden'), false); assert.equal(group.children[0].content, 'new');
  const redo = doc.applyEditOperation(undo).inverse; assert.deepEqual(slide.elements[0], before); assert.equal(slide.hidden, false);
  doc.applyEditOperation(redo); assert.equal(slide.elements[0].children[0].opacity, 0);
});

test('element property conflicts and invalid values reject complete mixed batches', () => {
  const doc = fixture(), slide = doc.slides[0]; slide.addElement({ id: 'text', type: 'text', content: 'value', hidden: false, opacity: 1 });
  const before = structuredClone(slide.elements), property = (type, expected, value) => ({ type, slideId: 'a', elementId: 'text', expected, value });
  for (const [bad, expected] of [[property('element-hidden', null, true), 'edit-conflict'], [property('element-opacity', null, 0), 'edit-conflict'],
    [property('element-opacity', 1, -0.1), 'invalid-operation'], [property('element-opacity', 1, 1.1), 'invalid-operation'],
    [property('element-opacity', 1, NaN), 'invalid-operation'], [property('element-hidden', false, 0), 'invalid-operation']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(expected));
    assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  }
  assert.throws(() => doc.applyEditOperation(op([property('element-opacity', 1, 0), property('element-opacity', 1, null)])), code('invalid-operation'));
  assert.throws(() => doc.applyEditOperation(op([{ ...property('element-hidden', false, true), elementId: 'missing' }])), code('missing-target'));
});

test('element names preserve absent and empty values across atomic undo and redo', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'group', type: 'group', name: '', box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [
    { id: 'child', type: 'text', content: 'old' },
  ] });
  const before = structuredClone(slide.elements), name = (id, expected, value) => ({ type: 'element-name', slideId: 'a', elementId: id, expected, value });
  const changes = [name('group', '', null), name('child', null, 'New & name'), { type: 'text-string', slideId: 'a', elementId: 'child', expected: 'old', value: 'new' }, hidden('a', false, true)];
  for (const [bad, error] of [[name('child', '', 'x'), 'edit-conflict'], [name('child', null, 1), 'invalid-operation'], [name('missing', null, 'x'), 'missing-target']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(error));
    assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  }
  assert.throws(() => doc.applyEditOperation(op([name('child', null, 'x'), name('child', null, 'y')])), code('invalid-operation'));
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse;
  assert.equal(Object.hasOwn(slide.elements[0], 'name'), false);
  assert.equal(slide.elements[0].children[0].name, 'New & name');
  const redo = doc.applyEditOperation(undo).inverse;
  assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  doc.applyEditOperation(redo); assert.equal(slide.elements[0].children[0].content, 'new');
});

test('transform batches preserve absent, empty and explicit defaults through undo and redo', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'group', type: 'group', transform: {}, box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [
    { id: 'child', type: 'text', content: 'old', transform: { rotation: 0, flipH: false } },
    { id: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 } },
  ] });
  const initial = structuredClone(slide.elements), change = (id, expected, value) => ({ type: 'element-transform', slideId: 'a', elementId: id, expected, value });
  const empty = { rotation: null, flipH: null, flipV: null }, initialChild = { rotation: 0, flipH: false, flipV: null };
  const changes = [change('group', empty, { rotation: 30, flipH: true, flipV: null }), change('child', initialChild, null), change('shape', null, empty),
    { type: 'text-string', slideId: 'a', elementId: 'child', expected: 'old', value: 'new' }, hidden('a', false, true)];
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse;
  assert.deepEqual(slide.elements[0].transform, { rotation: 30, flipH: true });
  assert.equal(Object.hasOwn(slide.elements[0].children[0], 'transform'), false);
  assert.deepEqual(slide.elements[0].children[1].transform, {});
  const redo = doc.applyEditOperation(JSON.parse(JSON.stringify(undo))).inverse;
  assert.deepEqual(slide.elements, initial); assert.equal(slide.hidden, false);
  doc.applyEditOperation(redo); assert.equal(slide.elements[0].children[0].content, 'new');
});

test('invalid, stale and ambiguous transform changes leave mixed batches unchanged', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 }, transform: { rotation: 0 } });
  const before = structuredClone(slide.elements), valid = { rotation: 0, flipH: null, flipV: null }, change = (expected, value) => ({ type: 'element-transform', slideId: 'a', elementId: 'shape', expected, value });
  for (const [bad, error] of [[change(null, valid), 'edit-conflict'], [change({ ...valid, flipH: false }, valid), 'edit-conflict'],
    [change(valid, { ...valid, rotation: Infinity }), 'invalid-operation'], [change(valid, { ...valid, flipV: 0 }), 'invalid-operation'],
    [change(valid, { ...valid, extra: 1 }), 'invalid-operation'], [change(valid, { rotation: 0 }), 'invalid-operation'],
    [{ ...change(valid, valid), elementId: 'missing' }, 'missing-target']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(error));
    assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  }
  let called = 0; const accessor = { ...valid }; Object.defineProperty(accessor, 'rotation', { enumerable: true, get() { called++; return 0; } });
  assert.throws(() => doc.applyEditOperation(op([change(valid, accessor)])), code('invalid-operation')); assert.equal(called, 0);
  assert.throws(() => doc.applyEditOperation(op([change(valid, null), change(valid, null)])), code('invalid-operation'));
  slide.addElement({ id: 'extended', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 }, transform: { rotation: 0, futureField: 'retain' } });
  const withExtension = structuredClone(slide.elements);
  assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), { ...change(valid, null), elementId: 'extended' }])), code('edit-conflict'));
  assert.deepEqual(slide.elements, withExtension); assert.equal(slide.hidden, false);
});

test('element actions replay atomically with nested text and distinct tooltip states', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'group', type: 'group', box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [
    { id: 'child', type: 'text', content: 'old', action: { type: 'url', url: 'https://example.org', tooltip: '' } },
    { id: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 } },
  ] });
  const before = structuredClone(slide.elements), action = (id, expected, value) => ({ type: 'element-action', slideId: 'a', elementId: id, expected, value });
  const old = { type: 'url', url: 'https://example.org', tooltip: '' }, link = { type: 'slide', slideId: 'b', tooltip: null };
  const changes = [action('child', old, link), action('shape', null, { type: 'url', url: 'mailto:test@example.org', tooltip: 'Mail' }),
    { type: 'text-string', slideId: 'a', elementId: 'child', expected: 'old', value: 'new' }, hidden('a', false, true)];
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse;
  assert.deepEqual(slide.elements[0].children[0].action, { type: 'slide', slideId: 'b' });
  const redo = doc.applyEditOperation(JSON.parse(JSON.stringify(undo))).inverse;
  assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  doc.applyEditOperation(redo); assert.equal(slide.elements[0].children[0].content, 'new');
});

test('action edits reject dangling destinations, stale values and invalid data without partial changes', () => {
  const doc = fixture(), slide = doc.slides[0]; slide.addElement({ id: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 } });
  const before = structuredClone(slide.elements), action = (expected, value) => ({ type: 'element-action', slideId: 'a', elementId: 'shape', expected, value });
  for (const [bad, error] of [[action(null, { type: 'slide', slideId: 'missing', tooltip: null }), 'missing-target'],
    [action({ type: 'url', url: 'https://example.org', tooltip: null }, null), 'edit-conflict'],
    [action(null, { type: 'url', url: ' ', tooltip: null }), 'invalid-operation'],
    [action(null, { type: 'url', url: 'x', tooltip: null, extra: 1 }), 'invalid-operation'],
    [action(null, { type: 'slide', slideId: 'b' }), 'invalid-operation'],
    [action(null, { type: 'other', tooltip: null }), 'invalid-operation']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(error));
    assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  }
  let called = 0; const value = { type: 'url', tooltip: null }; Object.defineProperty(value, 'url', { enumerable: true, get() { called++; return 'x'; } });
  assert.throws(() => doc.applyEditOperation(op([action(null, value)])), code('invalid-operation')); assert.equal(called, 0);
  assert.throws(() => doc.applyEditOperation(op([action(null, null), action(null, null)])), code('invalid-operation'));
  slide.addElement({ id: 'extended-action', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 }, action: { type: 'url', url: 'https://example.org', futureField: 'keep' } });
  const extended = structuredClone(slide.elements);
  assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), { ...action({ type: 'url', url: 'https://example.org', tooltip: null }, null), elementId: 'extended-action' }])), code('edit-conflict'));
  assert.deepEqual(slide.elements, extended); assert.equal(slide.hidden, false);
});

test('undo rejects an action destination removed after the original change', () => {
  const doc = fixture(), slide = doc.slides[0]; slide.addElement({ id: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 }, action: { type: 'slide', slideId: 'b' } });
  const undo = doc.applyEditOperation(op([{ type: 'element-action', slideId: 'a', elementId: 'shape', expected: { type: 'slide', slideId: 'b', tooltip: null }, value: null }])).inverse;
  doc.removeSlide('b');
  assert.throws(() => doc.applyEditOperation(undo), code('missing-target'));
  assert.equal(slide.elements[0].action, undefined);
});

test('run actions compose with text changes in shapes and table cells and replay independently', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'group', type: 'group', box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [
    { id: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 30, height: 20 }, text: { content: [{ runs: [{ text: 'shape', style: { bold: true } }] }] } },
    { id: 'table', type: 'table', box: { x: 0, y: 30, width: 50, height: 20 }, columns: [50], rows: [{ cells: [{ content: [{ runs: [{ text: 'cell', action: { type: 'url', url: 'https://example.org', tooltip: '' } }] }] }] }] },
  ] });
  const before = structuredClone(slide.elements), path = { slideId: 'a', elementId: 'shape', paragraphIndex: 0, runIndex: 0 }, cell = { slideId: 'a', elementId: 'table', rowIndex: 0, cellIndex: 0, paragraphIndex: 0, runIndex: 0 };
  const old = { type: 'url', url: 'https://example.org', tooltip: '' }, next = { type: 'slide', slideId: 'b', tooltip: null };
  const changes = [{ type: 'text-run', ...path, expected: 'shape', value: 'new shape' }, { type: 'text-run-action', ...path, expected: null, value: next },
    { type: 'table-cell-text-run', ...cell, expected: 'cell', value: 'new cell' }, { type: 'table-cell-text-run-action', ...cell, expected: old, value: null }];
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse;
  assert.deepEqual(slide.elements[0].children[0].text.content[0].runs[0].action, { type: 'slide', slideId: 'b' });
  assert.equal(slide.elements[0].children[1].rows[0].cells[0].content[0].runs[0].action, undefined);
  const redo = doc.applyEditOperation(undo).inverse; assert.deepEqual(slide.elements, before);
  doc.applyEditOperation(redo); assert.equal(slide.elements[0].children[1].rows[0].cells[0].content[0].runs[0].text, 'new cell');
});

test('run action failures reject all changes for missing paths, stale values and destinations', () => {
  const doc = fixture(), slide = doc.slides[0]; slide.addElement({ id: 'text', type: 'text', content: [{ runs: [{ text: 'x' }] }] });
  const before = structuredClone(slide.elements), change = { type: 'text-run-action', slideId: 'a', elementId: 'text', paragraphIndex: 0, runIndex: 0, expected: null, value: { type: 'slide', slideId: 'b', tooltip: null } };
  for (const [bad, error] of [[{ ...change, runIndex: 1 }, 'missing-target'], [{ ...change, runIndex: -1 }, 'invalid-operation'],
    [{ ...change, expected: { type: 'url', url: 'x', tooltip: null } }, 'edit-conflict'], [{ ...change, value: { type: 'slide', slideId: 'missing', tooltip: null } }, 'missing-target'],
    [{ ...change, value: { type: 'url', url: 'x' } }, 'invalid-operation']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(error));
    assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  }
  assert.throws(() => doc.applyEditOperation(op([change, change])), code('invalid-operation'));
  slide.addElement({ id: 'plain', type: 'text', content: 'x' });
  const withPlain = structuredClone(slide.elements);
  assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), { ...change, elementId: 'plain' }])), code('missing-target'));
  assert.deepEqual(slide.elements, withPlain); assert.equal(slide.hidden, false);
  const undo = doc.applyEditOperation(op([change])).inverse; doc.removeSlide('b');
  assert.throws(() => doc.applyEditOperation(op([change])), code('edit-conflict'));
  doc.applyEditOperation(undo); assert.equal(slide.elements[0].content[0].runs[0].action, undefined);
});

test('hiding an object restores authored reading order on inverse, including group scopes', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'first', type: 'text', content: 'first' });
  slide.addElement({ id: 'group', type: 'group', box: { x: 0, y: 0, width: 20, height: 20 }, coordinateSize: { width: 20, height: 20 }, children: [
    { id: 'one', type: 'text', content: 'one' }, { id: 'two', type: 'text', content: 'two' },
  ] });
  const original = [{ elementIds: ['group', 'first'] }, { groupId: 'group', elementIds: ['two', 'one'] }];
  slide.setReadingOrder(original);
  const request = op([{ type: 'element-hidden', slideId: 'a', elementId: 'group', expected: null, value: true }]);
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(request))).inverse;
  assert.deepEqual(slide.readingOrder, [{ elementIds: ['first'] }]);
  assert.equal(undo.changes.some(change => change.type === 'slide-reading-order'), true);
  const redo = doc.applyEditOperation(JSON.parse(JSON.stringify(undo))).inverse;
  assert.deepEqual(slide.readingOrder, original);
  doc.applyEditOperation(redo); assert.deepEqual(slide.readingOrder, [{ elementIds: ['first'] }]);
});

test('reading order changes are validated against resulting objects and conflict atomically', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'one', type: 'text', content: 'one' }); slide.addElement({ id: 'two', type: 'text', content: 'two' });
  const original = [{ elementIds: ['two', 'one'] }]; slide.setReadingOrder(original);
  const hide = { type: 'element-hidden', slideId: 'a', elementId: 'two', expected: null, value: true };
  const order = (expected, value) => ({ type: 'slide-reading-order', slideId: 'a', expected, value });
  for (const [bad, error] of [[order(null, null), 'edit-conflict'], [order(original, [{ elementIds: ['two', 'one'] }]), 'invalid-operation'],
    [order(original, []), 'invalid-operation'], [order(original, [{ elementIds: ['one', 'one'] }]), 'invalid-operation']]) {
    assert.throws(() => doc.applyEditOperation(op([hide, bad])), code(error));
    assert.deepEqual(slide.readingOrder, original); assert.equal(slide.elements[1].hidden, undefined);
  }
  const undo = doc.applyEditOperation(op([hide, order(original, [{ elementIds: ['one'] }])])).inverse;
  assert.deepEqual(slide.readingOrder, [{ elementIds: ['one'] }]);
  doc.applyEditOperation(undo); assert.deepEqual(slide.readingOrder, original); assert.equal(slide.elements[1].hidden, undefined);
  const change = order(original, null), inverse = doc.applyEditOperation(op([change])).inverse;
  assert.equal(slide.readingOrder, undefined); doc.applyEditOperation(inverse); assert.deepEqual(slide.readingOrder, original);
});

test('undo detects a changed reading order before restoring visibility', () => {
  const doc = fixture(), slide = doc.slides[0];
  for (const id of ['one', 'two', 'three']) slide.addElement({ id, type: 'text', content: id });
  slide.setReadingOrder([{ elementIds: ['three', 'two', 'one'] }]);
  const inverse = doc.applyEditOperation(op([{ type: 'element-hidden', slideId: 'a', elementId: 'three', expected: null, value: true }])).inverse;
  slide.setReadingOrder([{ elementIds: ['one', 'two'] }]);
  assert.throws(() => doc.applyEditOperation(inverse), code('edit-conflict'));
  assert.equal(slide.elements[2].hidden, true); assert.deepEqual(slide.readingOrder, [{ elementIds: ['one', 'two'] }]);
});

test('accessibility edits preserve empty and absent values and restore group reading order', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'first', type: 'text', content: 'first', accessibility: {} });
  slide.addElement({ id: 'group', type: 'group', box: { x: 0, y: 0, width: 20, height: 20 }, coordinateSize: { width: 20, height: 20 }, children: [
    { id: 'child', type: 'text', content: 'child', accessibility: { description: '', decorative: false } },
  ] });
  const original = [{ elementIds: ['group', 'first'] }, { groupId: 'group', elementIds: ['child'] }]; slide.setReadingOrder(original);
  const before = structuredClone(slide.elements), edit = (id, expected, value) => ({ type: 'element-accessibility', slideId: 'a', elementId: id, expected, value });
  const changes = [edit('group', null, { description: 'Artwork', decorative: true }),
    edit('first', { description: null, decorative: null }, null), edit('child', { description: '', decorative: false }, { description: 'Child', decorative: null })];
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse;
  assert.deepEqual(slide.readingOrder, [{ elementIds: ['first'] }]);
  assert.equal(Object.hasOwn(slide.elements[0], 'accessibility'), false);
  assert.deepEqual(slide.elements[1].children[0].accessibility, { description: 'Child' });
  const redo = doc.applyEditOperation(undo).inverse; assert.deepEqual(slide.elements, before); assert.deepEqual(slide.readingOrder, original);
  doc.applyEditOperation(redo); assert.deepEqual(slide.elements[1].accessibility, { description: 'Artwork', decorative: true });
});

test('accessibility changes reject stale, invalid and unsupported fields atomically', () => {
  const doc = fixture(), slide = doc.slides[0]; slide.addElement({ id: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 }, accessibility: { description: '' } });
  const before = structuredClone(slide.elements), valid = { description: '', decorative: null }, edit = (expected, value) => ({ type: 'element-accessibility', slideId: 'a', elementId: 'shape', expected, value });
  for (const [bad, error] of [[edit(null, valid), 'edit-conflict'], [edit({ description: '', decorative: false }, valid), 'edit-conflict'],
    [edit(valid, { description: 4, decorative: true }), 'invalid-operation'], [edit(valid, { description: null, decorative: 'yes' }), 'invalid-operation'],
    [edit(valid, { ...valid, future: 1 }), 'invalid-operation'], [{ ...edit(valid, null), elementId: 'missing' }, 'missing-target']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(error)); assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  }
  let calls = 0; const accessor = { decorative: null }; Object.defineProperty(accessor, 'description', { enumerable: true, get() { calls++; return ''; } });
  assert.throws(() => doc.applyEditOperation(op([edit(valid, accessor)])), code('invalid-operation')); assert.equal(calls, 0);
  assert.throws(() => doc.applyEditOperation(op([edit(valid, null), edit(valid, null)])), code('invalid-operation'));
  slide.addElement({ id: 'extended', type: 'text', content: 'x', accessibility: { description: 'x', future: 1 } });
  const extended = structuredClone(slide.elements);
  assert.throws(() => doc.applyEditOperation(op([{ ...edit({ description: 'x', decorative: null }, null), elementId: 'extended' }])), code('edit-conflict'));
  assert.deepEqual(slide.elements, extended);
});

test('notes run text and action edits preserve formatting and replay alongside element edits', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.setNotes([{ runs: [{ text: 'Notes', style: { bold: true }, action: { type: 'url', url: 'https://example.org', tooltip: '' } }, { text: ' tail' }] }]);
  slide.addElement({ id: 'item', type: 'text', content: 'old' });
  const before = structuredClone(slide.notes), request = op([
    { type: 'notes-run', slideId: 'a', paragraphIndex: 0, runIndex: 0, expected: 'Notes', value: 'Updated' },
    { type: 'notes-run-action', slideId: 'a', paragraphIndex: 0, runIndex: 0, expected: { type: 'url', url: 'https://example.org', tooltip: '' }, value: { type: 'slide', slideId: 'b', tooltip: null } },
    { type: 'text-string', slideId: 'a', elementId: 'item', expected: 'old', value: 'new' },
  ]);
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(request))).inverse;
  assert.deepEqual(slide.notes[0].runs[0], { text: 'Updated', style: { bold: true }, action: { type: 'slide', slideId: 'b' } });
  const redo = doc.applyEditOperation(undo).inverse; assert.deepEqual(slide.notes, before); assert.equal(slide.elements[0].content, 'old');
  doc.applyEditOperation(redo); assert.equal(slide.notes[0].runs[1].text, ' tail');
});

test('notes changes reject missing, stale and malformed paths without partial updates', () => {
  const doc = fixture(), slide = doc.slides[0]; slide.setNotes([{ runs: [{ text: 'one' }, { text: 'two' }] }]);
  const before = structuredClone(slide.notes), run = { type: 'notes-run', slideId: 'a', paragraphIndex: 0, runIndex: 0, expected: 'one', value: 'new' };
  for (const [bad, error] of [[{ ...run, expected: 'stale' }, 'edit-conflict'], [{ ...run, runIndex: 8 }, 'missing-target'],
    [{ ...run, paragraphIndex: -1 }, 'invalid-operation'], [{ type: 'notes-string', slideId: 'a', expected: 'one', value: 'x' }, 'missing-target'],
    [{ ...run, slideId: 'missing' }, 'missing-target'], [{ type: 'notes-run-action', slideId: 'a', paragraphIndex: 0, runIndex: 1, expected: null, value: { type: 'slide', slideId: 'missing', tooltip: null } }, 'missing-target']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(error)); assert.deepEqual(slide.notes, before); assert.equal(slide.hidden, false);
  }
  assert.throws(() => doc.applyEditOperation(op([run, run])), code('invalid-operation'));
  const undo = doc.applyEditOperation(op([run])).inverse; slide.setNotes([{ runs: [{ text: 'later' }] }]);
  assert.throws(() => doc.applyEditOperation(undo), code('edit-conflict')); assert.equal(slide.notes[0].runs[0].text, 'later');
});

test('notes string shorthand edits preserve representation and distinguish absent notes', () => {
  const doc = fixture(), slide = doc.slides[0], change = { type: 'notes-string', slideId: 'a', expected: '', value: 'hello' };
  assert.throws(() => doc.applyEditOperation(op([change])), code('missing-target'));
  slide.setNotes(''); const undo = doc.applyEditOperation(op([change])).inverse;
  assert.equal(slide.notes, 'hello'); const redo = doc.applyEditOperation(undo).inverse; assert.equal(slide.notes, '');
  doc.applyEditOperation(redo); assert.equal(slide.notes, 'hello');
});

test('chart point edits support categorical, combo, scatter and bubble data with atomic undo', () => {
  const doc = fixture(), slide = doc.slides[0], box = { x: 0, y: 0, width: 100, height: 80 };
  slide.addElement({ id: 'group', type: 'group', box, coordinateSize: { width: 100, height: 80 }, children: [
    { id: 'bar', type: 'chart', chartType: 'bar', box, categories: ['a', 'b'], series: [{ name: 's', values: [1, null] }] },
    { id: 'combo', type: 'chart', chartType: 'combo', box, categories: ['a'], series: [{ kind: 'bar', name: 'b', values: [2] }, { kind: 'line', name: 'l', values: [3] }] },
    { id: 'scatter', type: 'chart', chartType: 'scatter', box, series: [{ name: 's', xValues: [1, null], yValues: [2, 4] }] },
    { id: 'bubble', type: 'chart', chartType: 'bubble', box, series: [{ name: 's', xValues: [1], yValues: [2], sizes: [3] }] },
    { id: 'pie', type: 'chart', chartType: 'pie', box, categories: ['a'], series: [{ name: 's', values: [5] }] },
  ] });
  const before = structuredClone(slide.elements), point = (elementId, coordinate, seriesIndex, pointIndex, expected, value) => ({ type: 'chart-point', slideId: 'a', elementId, coordinate, seriesIndex, pointIndex, expected, value });
  const changes = [point('bar', 'values', 0, 1, null, 7), point('combo', 'values', 1, 0, 3, -3), point('scatter', 'xValues', 0, 1, null, 5),
    point('scatter', 'yValues', 0, 0, 2, null), point('bubble', 'sizes', 0, 0, 3, 8), point('pie', 'values', 0, 0, 5, 0), hidden('a', false, true)];
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse;
  assert.equal(slide.elements[0].children[1].series[1].values[0], -3); assert.equal(slide.elements[0].children[3].series[0].sizes[0], 8);
  const redo = doc.applyEditOperation(undo).inverse; assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  doc.applyEditOperation(redo); assert.equal(slide.elements[0].children[2].series[0].xValues[1], 5);
});

test('chart point changes reject bad paths, invalid values and stale values before any mutation', () => {
  const doc = fixture(), slide = doc.slides[0], box = { x: 0, y: 0, width: 100, height: 80 };
  slide.addElement({ id: 'bubble', type: 'chart', chartType: 'bubble', box, series: [{ name: 's', xValues: [1], yValues: [2], sizes: [3] }] });
  slide.addElement({ id: 'pie', type: 'chart', chartType: 'pie', box, categories: ['a'], series: [{ name: 's', values: [5] }] });
  const before = structuredClone(slide.elements), point = (elementId, coordinate, expected, value) => ({ type: 'chart-point', slideId: 'a', elementId, coordinate, seriesIndex: 0, pointIndex: 0, expected, value });
  for (const [bad, error] of [[point('bubble', 'sizes', 3, 0), 'invalid-operation'], [point('bubble', 'sizes', 3, null), 'invalid-operation'],
    [point('bubble', 'xValues', 1, null), 'invalid-operation'], [point('pie', 'values', 5, -1), 'invalid-operation'],
    [point('bubble', 'values', 3, 2), 'missing-target'], [point('bubble', 'sizes', 4, 5), 'edit-conflict'],
    [{ ...point('bubble', 'sizes', 3, 5), pointIndex: 2 }, 'missing-target'], [point('bubble', 'sizes', 3, NaN), 'invalid-operation']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(error)); assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  }
  const valid = point('bubble', 'sizes', 3, 4); assert.throws(() => doc.applyEditOperation(op([valid, valid])), code('invalid-operation'));
});

test('chart point negative zero has a JSON-stable token and distinct conflict behavior', () => {
  const doc = fixture(), slide = doc.slides[0];
  slide.addElement({ id: 'chart', type: 'chart', chartType: 'bar', box: { x: 0, y: 0, width: 100, height: 80 }, categories: ['a'], series: [{ name: 's', values: [-0] }] });
  const point = (expected, value) => ({ type: 'chart-point', slideId: 'a', elementId: 'chart', coordinate: 'values', seriesIndex: 0, pointIndex: 0, expected, value });
  assert.throws(() => doc.applyEditOperation(op([point(0, 1)])), code('edit-conflict'));
  assert.throws(() => doc.applyEditOperation(op([point(-0, 1)])), code('invalid-operation'));
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op([point('-0', 0)])))).inverse;
  assert.ok(Object.is(slide.elements[0].series[0].values[0], 0));
  const redo = doc.applyEditOperation(JSON.parse(JSON.stringify(undo))).inverse;
  assert.ok(Object.is(slide.elements[0].series[0].values[0], -0));
  doc.applyEditOperation(JSON.parse(JSON.stringify(redo))); assert.ok(Object.is(slide.elements[0].series[0].values[0], 0));
});

test('chart labels and numeric points compose with atomic undo across chart families', () => {
  const doc = fixture(), slide = doc.slides[0], box = { x: 0, y: 0, width: 100, height: 80 };
  slide.addElement({ id: 'group', type: 'group', box, coordinateSize: { width: 100, height: 80 }, children: [
    { id: 'bar', type: 'chart', chartType: 'bar', box, categories: ['Q1', 'Q2'], series: [{ name: 'Sales', values: [2, 4] }] },
    { id: 'scatter', type: 'chart', chartType: 'scatter', box, series: [{ name: 'Points', xValues: [1], yValues: [2] }] },
  ] });
  const before = structuredClone(slide.elements), target = { slideId: 'a', elementId: 'bar' };
  const changes = [{ type: 'chart-category', ...target, categoryIndex: 1, expected: 'Q2', value: '二 & two' },
    { type: 'chart-series-name', ...target, seriesIndex: 0, expected: 'Sales', value: 'Revenue' },
    { type: 'chart-point', ...target, seriesIndex: 0, pointIndex: 1, coordinate: 'values', expected: 4, value: 8 },
    { type: 'chart-series-name', slideId: 'a', elementId: 'scatter', seriesIndex: 0, expected: 'Points', value: '' }, hidden('a', false, true)];
  const undo = doc.applyEditOperation(JSON.parse(JSON.stringify(op(changes)))).inverse;
  assert.equal(slide.elements[0].children[0].categories[1], '二 & two'); assert.equal(slide.elements[0].children[1].series[0].name, '');
  const redo = doc.applyEditOperation(undo).inverse; assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  doc.applyEditOperation(redo); assert.equal(slide.elements[0].children[0].series[0].values[1], 8);
});

test('chart label changes reject stale, missing and malformed targets atomically', () => {
  const doc = fixture(), slide = doc.slides[0], box = { x: 0, y: 0, width: 100, height: 80 };
  slide.addElement({ id: 'bar', type: 'chart', chartType: 'bar', box, categories: ['A'], series: [{ name: 'S', values: [1] }] });
  slide.addElement({ id: 'bubble', type: 'chart', chartType: 'bubble', box, series: [{ name: 'B', xValues: [1], yValues: [2], sizes: [3] }] });
  const before = structuredClone(slide.elements), category = { type: 'chart-category', slideId: 'a', elementId: 'bar', categoryIndex: 0, expected: 'A', value: 'B' };
  for (const [bad, error] of [[{ ...category, expected: 'wrong' }, 'edit-conflict'], [{ ...category, categoryIndex: 3 }, 'missing-target'],
    [{ ...category, elementId: 'bubble' }, 'missing-target'], [{ ...category, categoryIndex: -1 }, 'invalid-operation'],
    [{ ...category, value: 3 }, 'invalid-operation'], [{ type: 'chart-series-name', slideId: 'a', elementId: 'bar', seriesIndex: 3, expected: '', value: 'x' }, 'missing-target']]) {
    assert.throws(() => doc.applyEditOperation(op([hidden('a', false, true), bad])), code(error)); assert.deepEqual(slide.elements, before); assert.equal(slide.hidden, false);
  }
  assert.throws(() => doc.applyEditOperation(op([category, category])), code('invalid-operation'));
});
