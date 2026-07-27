"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.paginate = void 0;
// serverside/src/utils/paginate.ts
const paginate = (query, page = 1, limit = 20) => {
    const skip = (Math.max(1, page) - 1) * Math.min(100, limit);
    return query.skip(skip).limit(Math.min(100, limit));
};
exports.paginate = paginate;
