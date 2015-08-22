using Xunit;

namespace Vba.Language.Tests.Compiler
{
    public class FunctionDeclarationTests
    {
        [Fact]
        public void CanParseImplicitFunctionDefinition()
        {
            var source = "Function MyFunction()\r\nEnd Function\r\n";
            var regex = @"(\(functionDeclaration.*\(identifier MyFunction\).*End Function.*\))";

            ParserHelper.CanParseSourceRegex(source, regex, p => p.functionDeclaration());
        }

        [Fact]
        public void CanParsePublicFunctionDefinition()
        {
            var source = "Public Function MyFunction()\r\nEnd Function\r\n";
            var regex = @"(\(functionDeclaration.*Public.*\(identifier MyFunction\).*End Function.*\))";

            ParserHelper.CanParseSourceRegex(source, regex, p => p.functionDeclaration());
        }

        [Fact]
        public void CanParsePrivateFunctionDefinition()
        {
            var source = "Private Function MyFunction()\r\nEnd Function\r\n";
            var regex = @"(\(functionDeclaration.*Private.*\(identifier MyFunction\).*End Function.*\))";

            ParserHelper.CanParseSourceRegex(source, regex, p => p.functionDeclaration());
        }
    }
}
