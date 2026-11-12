/** Shared public-API edit sequence; self-contained for browser and installed consumers. */
export function editMasterDefinitions(document) {
  document.duplicateSlide('page-0', document.slides.length);
  document.defineSlideMaster({ id: 'green', name: 'Green', theme: { colors: { accent1: '118855' }, fonts: { body: 'Verdana' } }, background: { type: 'solid', color: 'EEFFEE' } });
  const placeholders = [{ key: 'same', kind: 'body', box: { x: 30, y: 40, width: 300, height: 60 }, textStyle: { run: { bold: false, italic: false }, paragraph: { spaceBefore: 0 } } }];
  document.replaceSlideLayout('layout-0', { name: 'Shared green', masterId: 'green', placeholders, elements: [{ id: 'shared-mark', type: 'shape', shape: 'rect', box: { x: 350, y: 20, width: 20, height: 20 }, style: { fill: { type: 'solid', color: { theme: 'accent1' } } } }] });
  document.replaceSlideLayout('layout-2', { name: 'Default owner with no fill', placeholders, background: { type: 'none' } });
  document.replaceSlideMaster('red', { name: 'Red now inherits' });
  const prior = document.layouts[0];
  let rejected = false;
  try { document.replaceSlideLayout('layout-0', { name: 'Would break shared placeholders' }); }
  catch (error) { rejected = error.name === 'PresentationValidationError' && error.diagnostics.some(d => d.code === 'missing-placeholder'); }
  if (!rejected || document.layouts[0] !== prior) throw new Error('Shared replacement failed to roll back');
  return document;
}
