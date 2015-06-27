using Xunit;

namespace Vba.Language.Tests.Preprocessor
{
    public class ConditionalTests
    {
        [Theory]
        [InlineData("#If True Then", "(ifStatement # If (expression (boolExpression (boolLiteral True))) Then)")]
        [InlineData("#If False Then", "(ifStatement # If (expression (boolExpression (boolLiteral False))) Then)")]
        [InlineData("#If abc = True Then", "(ifStatement # If (expression (expression abc) (comparisonOperator =) (expression (boolExpression (boolLiteral True)))) Then)")]
        [InlineData("#If abc = False Then", "(ifStatement # If (expression (expression abc) (comparisonOperator =) (expression (boolExpression (boolLiteral False)))) Then)")]
        public void CanParseValidIfStatement(string source, string expectedTree)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.ifStatement();

            Assert.Null(result.exception);
            ParseTreeHelper.TreesAreEqual(expectedTree, result.ToStringTree(parser));
        }

        [Theory]
        [InlineData("#ElseIf True Then", "(elseIfStatement # ElseIf (expression (boolExpression (boolLiteral True))) Then)")]
        [InlineData("#ElseIf False Then", "(elseIfStatement # ElseIf (expression (boolExpression (boolLiteral False))) Then)")]
        [InlineData("#ElseIf abc = True Then", "(elseIfStatement # ElseIf (expression (expression abc) (comparisonOperator =) (expression (boolExpression (boolLiteral True)))) Then)")]
        [InlineData("#ElseIf abc = False Then", "(elseIfStatement # ElseIf (expression (expression abc) (comparisonOperator =) (expression (boolExpression (boolLiteral False)))) Then)")]
        public void CanParseValidElseIfStatement(string source, string expectedTree)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.elseIfStatement();

            Assert.Null(result.exception);
            ParseTreeHelper.TreesAreEqual(expectedTree, result.ToStringTree(parser));
        }

        [Theory]
        [InlineData("#Else", "(elseStatement # Else)")]
        [InlineData("   #Else", "(elseStatement # Else)")]
        [InlineData("   #Else   ", "(elseStatement # Else)")]
        [InlineData("\t\t#Else", "(elseStatement # Else)")]
        [InlineData("#Else\t\t", "(elseStatement # Else)")]
        [InlineData("\t\t#Else\t\t", "(elseStatement # Else)")]
        public void CanParseValidElseStatement(string source, string expectedTree)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.elseStatement();

            Assert.Null(result.exception);
            ParseTreeHelper.TreesAreEqual(expectedTree, result.ToStringTree(parser));
        }

        [Theory]
        [InlineData("#End If", "(endIfStatement # End If)")]
        [InlineData("   #End If", "(endIfStatement # End If)")]
        [InlineData("   #End If   ", "(endIfStatement # End If)")]
        [InlineData("\t\t#End If", "(endIfStatement # End If)")]
        [InlineData("#End If\t\t", "(endIfStatement # End If)")]
        [InlineData("\t\t#End If\t\t", "(endIfStatement # End If)")]
        public void CanParseValidEndIfStatement(string source, string expectedTree)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.endIfStatement();

            Assert.NotNull(result.End());
            Assert.NotNull(result.If());
            Assert.Null(result.exception);
        }
    }
}
