using System.Linq;
using Vba.Grammars;
using Xunit;

namespace Vba.Language.Tests.Compiler
{
    public class LexerTests
    {
        [Theory]
        [InlineData("#01/02/2000#")]
        public void CanLexValidDate(string source)
        {
            AssertTokenIsType(source, VbaLexer.DateLiteral);
        }

        [Theory]
        [InlineData("\"\"")]
        [InlineData("\"Random Text\"")]
        [InlineData("\"\"\"\"")]
        [InlineData("\"Leading text \"\" Trailing Text\"")]
        public void CanLexValidString(string source)
        {
            AssertTokenIsType(source, VbaLexer.StringLiteral);
        }

        [Theory]
        [InlineData("0")]
        [InlineData("1")]
        [InlineData("1234567890")]
        [InlineData("&O1")]
        [InlineData("&O1234")]
        [InlineData("&O5670")]
        [InlineData("&H1")]
        [InlineData("&H12345")]
        [InlineData("&H67890")]
        [InlineData("&HABCDEF")]
        public void CanLexValidInteger(string source)
        {
            AssertTokenIsType(source, VbaLexer.IntegerLiteral);
        }

        [Theory]
        [InlineData("0.0")]
        [InlineData("12345.67890")]
        [InlineData("12345.67890!")]
        [InlineData("12345.67890#")]
        [InlineData("12345.67890@")]
        [InlineData("1234567890!")]
        [InlineData("1234567890#")]
        [InlineData("1234567890@")]
        public void CanLexValidFloat(string source)
        {
            AssertTokenIsType(source, VbaLexer.FloatLiteral);
        }

        private static void AssertTokenIsType(string source, int tokenType)
        {
            var lexer = VbaCompilerHelper.BuildVbaLexer(source);

            var tokens = lexer.GetAllTokens();

            Assert.Equal(1, tokens.Count);
            Assert.Equal(tokenType, tokens[0].Type);
        }
    }
}
