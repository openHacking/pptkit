let presentationCounter = 0;

export function createDocumentId(): string {
  presentationCounter += 1;
  return `presentation-${presentationCounter}`;
}

export class IdAllocator {
  private counter = 0;

  constructor(private readonly prefix: string) {}

  fork(): IdAllocator {
    const allocator = new IdAllocator(this.prefix);
    allocator.counter = this.counter;
    return allocator;
  }

  next(isUsed: (id: string) => boolean): string {
    let id: string;
    do {
      this.counter += 1;
      id = `${this.prefix}-${this.counter}`;
    } while (isUsed(id));
    return id;
  }
}
