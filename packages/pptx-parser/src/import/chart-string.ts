/** ST_Xstring uses one-pass UTF-16 escapes; an escaped underscore is not decoded twice. */
export function chartString(value: string): string {
  return value.replace(/_x([0-9a-fA-F]{4})_/g, (_match, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)));
}
