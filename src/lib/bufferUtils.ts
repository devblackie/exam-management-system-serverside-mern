// src/lib/bufferUtils.ts
export function toNodeBuffer(buf: unknown): Buffer {
  // if it's already a Node Buffer, return it
  if (Buffer.isBuffer(buf)) return buf as Buffer;

  // if it's a Uint8Array / ArrayBuffer-like, convert
  try {
    // Handle ArrayBuffer, Uint8Array, Buffer-like
    if (buf instanceof ArrayBuffer) return Buffer.from(new Uint8Array(buf));
    if (ArrayBuffer.isView(buf)) return Buffer.from(buf as Uint8Array);
    // fallback: attempt to coerce (unsafe, but will often work)
    return Buffer.from(buf as any);
  } catch (err) {
    // final fallback
    return Buffer.from([]);
  }
}
