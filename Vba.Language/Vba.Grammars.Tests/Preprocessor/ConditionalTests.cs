using Xunit;

namespace Vba.Grammars.Tests.Preprocessor
{
    public class ConditionalTests
    {
        [Theory]
        [InlineData("#If True Then", "True")]
        [InlineData("#If False Then", "False")]
        [InlineData("#If abc = True Then", "abc = True")]
        [InlineData("#If abc = False Then", "abc = False")]
        public void CanParseValidIfStatement(string source, string expression)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.ifStatement();

            Assert.Equal(expression, result.expression().GetText());
            Assert.Null(result.exception);
        }
    }
}
