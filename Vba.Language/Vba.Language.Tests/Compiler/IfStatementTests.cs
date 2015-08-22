using Xunit;

namespace Vba.Language.Tests.Compiler
{
    public class IfStatementTests
    {
        [Fact]
        public void CanParseSimpleMultilineIfStatement()
        {
            var source = "If True Then\r\nEnd If\r\n";
            var regex = @"(\(ifStatement.*True.*\))";

            ParserHelper.CanParseSourceRegex(source, regex, p => p.statementBlock());
        }
    }
}
