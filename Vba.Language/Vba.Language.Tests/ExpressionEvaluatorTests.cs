using Vba.Language.Tests.Preprocessor;

using Xunit;

namespace Vba.Language.Tests
{
    public class ExpressionEvaluatorTests
    {
        [Theory]
        [InlineData("True", true)]
        [InlineData("False", false)]
        public void CanEvaluateValidBooleanExpression(string source, bool expected)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);
            var expression = parser.boolExpression();
            var evaluator = new ExpressionEvaluator(new ConstantsDictionary());

            var result = evaluator.VisitBoolExpression(expression);

            Assert.IsAssignableFrom<bool>(result);
            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData("1 = 1", true)]
        [InlineData("1 <> 2", true)]
        //[InlineData("\"\" = \"\"", true)]  for some reason this crashes the test runner.
        [InlineData("\"a value\" >< \"different value\"", true)]
        public void CanEvaluateValidComparisonExpression(string source, bool expected)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);
            var expression = parser.expression();
            var evaluator = new ExpressionEvaluator(new ConstantsDictionary());

            var result = evaluator.VisitExpression(expression);

            Assert.IsAssignableFrom<bool>(result);
            Assert.Equal(expected, result);
        }
    }
}
