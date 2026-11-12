import { SvgRenderAbortError } from '../types/public.js';

export function checkRenderAbort(signal?: AbortSignal): void {
  if (signal !== undefined && (signal === null || typeof signal !== 'object' || typeof signal.aborted !== 'boolean' || typeof signal.addEventListener !== 'function' || typeof signal.removeEventListener !== 'function')) throw new TypeError('signal must be an AbortSignal.');
  if (signal?.aborted) throw new SvgRenderAbortError();
}

/** Stop waiting promptly; the resolver owns its own underlying work and receives the signal. */
export async function waitForResolver<T>(promise: PromiseLike<T> | T, signal?: AbortSignal): Promise<T> {
  if (!signal) return await promise;
  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const finish = (callback: () => void) => { if (settled) return; settled = true; signal.removeEventListener('abort', abort); callback(); };
    const abort = () => finish(() => reject(new SvgRenderAbortError()));
    signal.addEventListener('abort', abort, { once: true });
    if (signal.aborted) abort();
    Promise.resolve(promise).then(value => finish(() => resolve(value)), error => finish(() => reject(error)));
  });
}
