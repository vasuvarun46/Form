export async function withRetry<T>(
  fn: () => Promise<T>,
  maxRetries: number = 3,
  delayMs: number = 1000
): Promise<T> {
  let lastError: Error | undefined;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await fn();
    } catch (e: unknown) {
      lastError = e instanceof Error ? e : new Error(String(e));

      // Only retry on transient errors (429 throttled, 503 service unavailable)
      const msg = lastError.message || "";
      const isTransient =
        msg.indexOf("429") >= 0 ||
        msg.indexOf("503") >= 0 ||
        msg.indexOf("throttled") >= 0;

      if (!isTransient || attempt === maxRetries) {
        throw lastError;
      }

      // Exponential backoff: 1s, 2s, 4s
      const wait = delayMs * Math.pow(2, attempt);
      await new Promise<void>((resolve) => setTimeout(resolve, wait));
    }
  }

  throw lastError || new Error("Retry failed");
}