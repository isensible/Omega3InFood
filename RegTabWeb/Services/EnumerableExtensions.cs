using System.Collections.Generic;

namespace RegTabWeb.Services
{
    public static class EnumerableExtensions
    {
        public static IEnumerable<T> SkipLastN<T>(this IEnumerable<T> source, int n) {
            // ReSharper disable once GenericEnumeratorNotDisposed
            var  it = source.GetEnumerator();
            bool hasRemainingItems = false;
            var  cache = new Queue<T>(n + 1);

            do {
                // ReSharper disable once AssignmentInConditionalExpression
                if (hasRemainingItems = it.MoveNext()) {
                    cache.Enqueue(it.Current);
                    if (cache.Count > n)
                        yield return cache.Dequeue();
                }
            } while (hasRemainingItems);
        }
    }
}