using Xunit;

namespace Vba.Language.Tests.Preprocessor
{
    public class ConstantTests
    {
        [Theory]
        [InlineData("#Const abc = True", "(constantDeclaration # Const abc = (expression (boolExpression (boolLiteral True))))")]
        [InlineData("#Const abc = \"\"", "(constantDeclaration # Const abc = (expression (stringExpression \"\")))")]
        [InlineData("#Const abc = 204", "(constantDeclaration # Const abc = (expression (arithmeticExpression 204)))")]
        [InlineData("#Const abc = -100.56", "(constantDeclaration # Const abc = (expression (arithmeticExpression -100.56)))")]
        public void CanParseValidConstantDefinition(string source, string expectedTree)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.constantDeclaration();

            Assert.Null(result.exception);
            ParseTreeHelper.TreesAreEqual(expectedTree, result.ToStringTree(parser));
        }
    }
}
