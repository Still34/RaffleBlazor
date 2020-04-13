using System;
using System.Collections.Generic;

namespace RaffleBlazor.Utility
{
    public class EnumerableUtility
    {
        private readonly Random _random;

        public EnumerableUtility(Random random)
            => _random = random;

        /// <seealso href="https://stackoverflow.com/a/1262619" />
        public List<T> Shuffle<T>(List<T> list)
        {
            var n = list.Count;
            while (n > 1)
            {
                n--;
                var k = _random.Next(n + 1);
                var value = list[k];
                list[k] = list[n];
                list[n] = value;
            }

            return list;
        }
    }
}