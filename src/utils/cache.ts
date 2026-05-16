// serverside/src/utils/cache.ts


import NodeCache from "node-cache";

const cache = new NodeCache({ stdTTL: 300, checkperiod: 60 });


export const cached = async <T>(
  key: string,
  fetcher: () => Promise<T>,
  ttlSeconds: number = 300,
): Promise<T> => {
  const hit = cache.get<T>(key);
  if (hit !== undefined) return hit; // cache hit — skip DB entirely

  const value = await fetcher(); // cache miss — go to MongoDB
  cache.set(key, value, ttlSeconds);
  return value;
};

/**
 * invalidateCache(prefix)
 *
 * Clears all cache keys that start with the given prefix.
 * Call this immediately after any write that changes the cached data.
 *
 * Examples:
 *   invalidateCache(`programs:${institutionId}`)    // after createProgram
 *   invalidateCache(`settings:${institutionId}`)    // after saveSettings
 *   invalidateCache(`units:${institutionId}`)        // after createUnit / deleteUnit
 *
 * You MUST call this after writes or the UI will show stale data for up to TTL seconds.
 * This is the only manual step required.
 */
export const invalidateCache = (prefix: string): void => {
  const keys = cache.keys().filter((k) => k.startsWith(prefix));
  if (keys.length > 0) cache.del(keys);
};

/**
 * invalidateCacheExact(key)
 *
 * Clears exactly one key. Use when you know the exact key.
 */
export const invalidateCacheExact = (key: string): void => {
  cache.del(key);
};
