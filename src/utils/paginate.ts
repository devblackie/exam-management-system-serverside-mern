// serverside/src/utils/paginate.ts
export const paginate = <T>(
  query: any,
  page: number = 1,
  limit: number = 20,
) => {
  const skip = (Math.max(1, page) - 1) * Math.min(100, limit);
  return query.skip(skip).limit(Math.min(100, limit));
};
