export function checkAssetSize(size: number, maximum?: number): void {
  if (maximum !== undefined && size > maximum) throw new RangeError(`Asset exceeds maxAssetBytes (${maximum}).`);
}

/** Abort waiting without requiring a cooperative third-party resolver to settle. */
export function withAbort<T>(pending: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return pending;
  return new Promise<T>((resolve, reject) => {
    const aborted = () => { cleanup(); reject(signal.reason); };
    const cleanup = () => signal.removeEventListener("abort", aborted);
    pending.then(value => { cleanup(); resolve(value); }, error => { cleanup(); reject(error); });
    signal.addEventListener("abort", aborted, { once: true });
    if (signal.aborted) aborted();
  });
}
