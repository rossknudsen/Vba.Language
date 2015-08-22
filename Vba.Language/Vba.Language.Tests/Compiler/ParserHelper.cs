using System;
using System.Text.RegularExpressions;
using Antlr4.Runtime;
using Vba.Grammars;
using Xunit;

namespace Vba.Language.Tests.Compiler
{
    internal class ParserHelper
    {
        [Obsolete("Use CanParseSourceRegex instead.")]
        internal static void CanParseSource(string source, string expectedTree, Func<VbaParser, ParserRuleContext> rule)
        {
            var parser = VbaCompilerHelper.BuildVbaParser(source);

            var result = rule(parser);

            Assert.Null(result.exception);
            ParseTreeHelper.TreesAreEqual(expectedTree, result.ToStringTree(parser));
        }

        internal static void CanParseSourceRegex(string source, string regex, Func<VbaParser, ParserRuleContext> rule)
        {
            var parser = VbaCompilerHelper.BuildVbaParser(source);

            var result = rule(parser);

            Assert.Null(result.exception);
            Assert.True(Regex.IsMatch(result.ToStringTree(parser), regex));
        }
    }
}