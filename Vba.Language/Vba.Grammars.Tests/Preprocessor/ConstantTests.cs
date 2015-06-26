using Xunit;

namespace Vba.Grammars.Tests.Preprocessor
{
    public class ConstantTests
    {
        [Theory]
        [InlineData("#Const abc = True", "abc", "True")]
        [InlineData("#Const abc = \"\"", "abc", "\"\"")]
        [InlineData("#Const abc = 204", "abc", "204")]
        [InlineData("#Const abc = -100.56", "abc", "-100.56")]
        public void CanParseValidConstantDefinition(string source, string id, string expression)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.constantDeclaration();

            Assert.Equal(id, result.ID().GetText());
            Assert.Equal(expression, result.expression().GetText());
            Assert.Null(result.exception);
        }
    }
}
