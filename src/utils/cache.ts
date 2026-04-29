// serverside/src/utils/cache.ts
//
// WHY CACHING HERE SPECIFICALLY
// ──────────────────────────────────────────────────────────────────────────
// node-cache is a simple in-process TTL cache backed by a plain JavaScript
// Map. It runs inside the same Node.js process — no Redis required.
//
// What it's good for in this system:
//   - Programs list (changes once a month at most)
//   - Institution settings (changes rarely)
//   - AcademicYear list (changes once a year)
//   - Unit templates (extremely stable)
//
// What it's BAD for in this system:
//   - Student marks (change daily during exam season)
//   - Student status (changes on every promotion run)
//   - Billing invoices (must always be live)
//   - Disciplinary cases (must always be live — stale cache here is a liability)
//   - Anything that's user-specific and institution-scoped across
//     multiple server instances (see securityStore.ts — that needs Redis)
//
// HOW TTL WORKS
// ──────────────────────────────────────────────────────────────────────────
// stdTTL: 300 means each key expires after 5 minutes automatically.
// checkperiod: 60 means node-cache sweeps for expired keys every 60 seconds.
// You don't need to think about cleanup — the library handles it.
//
// INSTALL
// ──────────────────────────────────────────────────────────────────────────
//   npm install node-cache
//   npm install --save-dev @types/node-cache

import NodeCache from "node-cache";

const cache = new NodeCache({ stdTTL: 300, checkperiod: 60 });

/**
 * cached<T>(key, fetcher, ttlSeconds?)
 *
 * If key exists in cache → return it immediately (zero DB calls).
 * If not → call fetcher(), store the result, return it.
 *
 * Generic T means TypeScript knows exactly what type comes back — no `any`.
 *
 * Usage examples:
 *
 *   // Cache programs for 10 minutes (they change rarely)
 *   const programs = await cached(
 *     `programs:${institutionId}`,
 *     () => Program.find({ institution: institutionId }).lean(),
 *     600,
 *   );
 *
 *   // Cache institution settings for 5 minutes (default TTL)
 *   const settings = await cached(
 *     `settings:${institutionId}`,
 *     () => InstitutionSettings.findOne({ institution: institutionId }).lean(),
 *   );
 *
 *   // Cache unit templates for 15 minutes (they almost never change)
 *   const units = await cached(
 *     `units:${institutionId}`,
 *     () => Unit.find({ institution: institutionId }).lean(),
 *     900,
 *   );
 */
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
