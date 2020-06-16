using System;

namespace Xsheet.Extensions
{
    public static class FuncExtensions
    {
        public static Func<T, TReturn2> Compose<T, TReturn1, TReturn2>(this Func<TReturn1, TReturn2> f1, Func<T, TReturn1> f2)
        {
            return x => f1(f2(x));
        }
    }
}
