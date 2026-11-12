/** Bounded source identity profile; callers retain the original spelling. */
export function diagramModelIdentity(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  const lexical = value.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
  if (/^[+-]?\d+$/.test(lexical)) {
    const number = Number(lexical);
    if (Number.isSafeInteger(number) && number >= -2147483648 && number <= 2147483647) return `int:${number}`;
  }
  return /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/.test(lexical) ? `guid:${lexical}` : undefined;
}
