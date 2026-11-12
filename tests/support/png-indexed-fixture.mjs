export function indexedRows(width, height, depth, invalidRow = -1, invalidIndex = 2) {
  const rows = [];
  for (let y = 0; y < height; y++) {
    const row = Buffer.alloc(Math.ceil(width * depth / 8));
    for (let x = 0; x < width; x++) {
      const value = y === invalidRow && x === width - 1 ? invalidIndex : depth === 1 ? 0 : (x + y) % 2;
      row[Math.floor(x * depth / 8)] |= value << (8 - depth - (x * depth % 8));
    }
    const padding = (8 - (width * depth % 8)) % 8;
    if (padding) row[row.length - 1] |= (1 << padding) - 1;
    rows.push(row);
  }
  return rows;
}

export function filterRows(rows) {
  const result = []; let previous = Buffer.alloc(rows[0].length);
  for (let y = 0; y < rows.length; y++) {
    const row = rows[y], filter = y % 5, encoded = Buffer.alloc(row.length + 1); encoded[0] = filter;
    for (let x = 0; x < row.length; x++) {
      const left = x ? row[x - 1] : 0, up = previous[x], upperLeft = x ? previous[x - 1] : 0;
      let predictor = 0;
      if (filter === 1) predictor = left;
      else if (filter === 2) predictor = up;
      else if (filter === 3) predictor = Math.floor((left + up) / 2);
      else if (filter === 4) { const p = left + up - upperLeft, a = Math.abs(p - left), b = Math.abs(p - up), c = Math.abs(p - upperLeft); predictor = a <= b && a <= c ? left : b <= c ? up : upperLeft; }
      encoded[x + 1] = (row[x] - predictor + 256) & 255;
    }
    result.push(encoded); previous = row;
  }
  return Buffer.concat(result);
}
