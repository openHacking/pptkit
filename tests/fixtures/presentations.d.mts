export const fixtureCatalog: readonly { id: string; capabilities: readonly string[] }[];
export function createFixture<T>(id: string, factory: (options: { id: string }) => T): T;
