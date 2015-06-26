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

        [Theory]
        [InlineData("#ElseIf True Then", "True")]
        [InlineData("#ElseIf False Then", "False")]
        [InlineData("#ElseIf abc = True Then", "abc = True")]
        [InlineData("#ElseIf abc = False Then", "abc = False")]
        public void CanParseValidElseIfStatement(string source, string expression)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.elseIfStatement();

            Assert.Equal(expression, result.expression().GetText());
            Assert.Null(result.exception);
        }

        [Theory]
        [InlineData("#Else")]
        [InlineData("   #Else")]
        [InlineData("   #Else   ")]
        [InlineData("\t\t#Else")]
        [InlineData("#Else\t\t")]
        [InlineData("\t\t#Else\t\t")]
        public void CanParseValidElseStatement(string source)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.elseStatement();

            Assert.NotNull(result.Else());
            Assert.Null(result.exception);
        }

        [Theory]
        [InlineData("#End If")]
        [InlineData("   #End If")]
        [InlineData("   #End If   ")]
        [InlineData("\t\t#End If")]
        [InlineData("#End If\t\t")]
        [InlineData("\t\t#End If\t\t")]
        public void CanParseValidEndIfStatement(string source)
        {
            var parser = PreprocessorHelper.BuildPreprocessorParser(source);

            var result = parser.endIfStatement();

            Assert.NotNull(result.End());
            Assert.NotNull(result.If());
            Assert.Null(result.exception);
        }
    }
}
