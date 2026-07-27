"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.toNodeBuffer = toNodeBuffer;
// src/lib/bufferUtils.ts
function toNodeBuffer(buf) {
    // if it's already a Node Buffer, return it
    if (Buffer.isBuffer(buf))
        return buf;
    // if it's a Uint8Array / ArrayBuffer-like, convert
    try {
        // Handle ArrayBuffer, Uint8Array, Buffer-like
        if (buf instanceof ArrayBuffer)
            return Buffer.from(new Uint8Array(buf));
        if (ArrayBuffer.isView(buf))
            return Buffer.from(buf);
        // fallback: attempt to coerce (unsafe, but will often work)
        return Buffer.from(buf);
    }
    catch (err) {
        // final fallback
        return Buffer.from([]);
    }
}
