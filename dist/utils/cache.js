"use strict";
// serverside/src/utils/cache.ts
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.invalidateCacheExact = exports.invalidateCache = exports.cached = void 0;
const node_cache_1 = __importDefault(require("node-cache"));
const cache = new node_cache_1.default({ stdTTL: 300, checkperiod: 60 });
const cached = async (key, fetcher, ttlSeconds = 300) => {
    const hit = cache.get(key);
    if (hit !== undefined)
        return hit; // cache hit — skip DB entirely
    const value = await fetcher(); // cache miss — go to MongoDB
    cache.set(key, value, ttlSeconds);
    return value;
};
exports.cached = cached;
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
const invalidateCache = (prefix) => {
    const keys = cache.keys().filter((k) => k.startsWith(prefix));
    if (keys.length > 0)
        cache.del(keys);
};
exports.invalidateCache = invalidateCache;
/**
 * invalidateCacheExact(key)
 *
 * Clears exactly one key. Use when you know the exact key.
 */
const invalidateCacheExact = (key) => {
    cache.del(key);
};
exports.invalidateCacheExact = invalidateCacheExact;
