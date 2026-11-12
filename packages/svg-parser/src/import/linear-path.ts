export interface PathPoint { x: number; y: number }

// Single-subpath move/line grammar, including implicit coordinate repetitions.
export function linearPath(data: string | undefined): { points: PathPoint[]; closed: boolean } | undefined {
  if (!data || data.length > 4096) return undefined;
  const number = '[+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][+-]?\\d+)?';
  const pair = new RegExp(`^[ \\t\\r\\n]*(${number})(?:[ \\t\\r\\n]*,[ \\t\\r\\n]*|[ \\t\\r\\n]+|(?=[+-]))(${number})`);
  const single = new RegExp(`^[ \\t\\r\\n]*(${number})`);
  let remaining = data.trim(), current: PathPoint = { x: 0, y: 0 }, closed = false, active = '';
  const points: PathPoint[] = [];
  while (remaining) {
    if (closed || points.length > 256) return undefined;
    const whitespace = /^[ \t\r\n]+/.exec(remaining)?.[0].length ?? 0;
    remaining = remaining.slice(whitespace);
    let command = remaining[0] ?? '';
    if (/[MmLlHhVvZz]/.test(command) && command) {
      remaining = remaining.slice(1);
      if (points.length && /[Mm]/.test(command)) return undefined;
      active = command;
    } else {
      if (!active || /[Zz]/.test(active)) return undefined;
      if (remaining.startsWith(',')) remaining = remaining.slice(1);
      else if (!whitespace && !/^[+-]/.test(remaining)) return undefined;
      command = active;
    }
    if (command === 'Z' || command === 'z') {
      if (points.length < 3) return undefined;
      points.push({ ...points[0]! }); closed = true;
    } else {
      if (!points.length && command !== 'M' && command !== 'm') return undefined;
      const match = (/[HhVv]/.test(command) ? single : pair).exec(remaining);
      if (!match) return undefined;
      const values = match.slice(1).map(Number);
      if (values.some(value => !Number.isFinite(value) || Math.abs(value) > 1_000_000)) return undefined;
      const relative = command === command.toLowerCase();
      const x = /[Vv]/.test(command) ? current.x : values[0]! + (relative ? current.x : 0);
      const y = /[Hh]/.test(command) ? current.y : (/[Vv]/.test(command) ? values[0]! : values[1]!) + (relative ? current.y : 0);
      if (!Number.isFinite(x) || !Number.isFinite(y) || Math.abs(x) > 1_000_000 || Math.abs(y) > 1_000_000) return undefined;
      current = { x, y }; points.push(current); remaining = remaining.slice(match[0].length);
      if (command === 'M') active = 'L';
      if (command === 'm') active = 'l';
    }
    if (!remaining.trim()) remaining = '';
  }
  if (points.length < 2 || points.length > 257) return undefined;
  if (points.some((point, index) => index > 0 && point.x === points[index - 1]!.x && point.y === points[index - 1]!.y)) return undefined;
  return { points, closed };
}
