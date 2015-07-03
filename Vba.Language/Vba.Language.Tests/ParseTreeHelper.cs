using System.Collections.Generic;
using System.Linq;

using Xunit;

namespace Vba.Language.Tests
{
    internal static class ParseTreeHelper
    {
        internal static void TreesAreEqual(string expected, string actual)
        {
            if (expected == null || actual == null)
            {
                Assert.True(false, "Expected and/or Actual are null.");
            }
            var filteredExpected = RemoveWhiteSpace(expected);
            var filteredActual = RemoveWhiteSpace(actual);
            Assert.Equal(filteredExpected, filteredActual);
        }

        private static string RemoveWhiteSpace(string input)
        {
            // the final \\t replacement is necessary because antlr seems to add it to the ToStringTree method. 
            return input.Replace("\t", "").Replace(" ", "").Replace("\\t", "");
        }

        internal static IEnumerable<char> ConvertTabsToSpaces(this IEnumerable<char> input)
        {
            foreach (var c in input)
            {
                if (c =='\t')
                {
                    yield return ' ';
                }
                yield return c;
            }
        }

        internal static IEnumerable<char> RemoveDuplicateSpaces(this IEnumerable<char> input)
        {
            var prevCharIsSpace = false;
            foreach (var c in input)
            {
                if (!prevCharIsSpace)
                {
                    yield return c;
                }
                prevCharIsSpace = (c == ' ');
            }
        }
    }
}