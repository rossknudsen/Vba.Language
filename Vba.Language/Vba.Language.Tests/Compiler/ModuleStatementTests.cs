using Xunit;

namespace Vba.Language.Tests.Compiler
{
    public class ModuleStatementTests
    {
        [Theory]
        [InlineData("Option Compare Text", "(directiveElement (optionCompareDirective Option Compare Text))")]
        [InlineData("Option Compare Binary", "(directiveElement (optionCompareDirective Option Compare Binary))")]
        [InlineData("Option Compare Database", "(directiveElement (optionCompareDirective Option Compare Database))")]
        [InlineData("Option Base 0", "(directiveElement (optionBaseDirective Option Base 0))")]
        [InlineData("Option Base 1", "(directiveElement (optionBaseDirective Option Base 1))")]
        [InlineData("Option Explicit", "(directiveElement (optionExplicitDirective Option Explicit))")]
        [InlineData("Option Private Module", "(directiveElement (optionPrivateDirective Option Private Module))")]
        [InlineData("Implements InterfaceName", "(directiveElement (implementsDirective Implements InterfaceName))")]
        [InlineData("DefStr A", "(directiveElement (defDirective (defType DefStr) (letterSpec A)))")]
        [InlineData("DefStr A-Q", "(directiveElement (defDirective (defType DefStr) (letterSpec A - Q)))")]
        public void CanParseDirectiveElement(string source, string expectedTree)
        {
            var parser = VbaCompilerHelper.BuildPreprocessorParser(source);

            var result = parser.directiveElement();

            Assert.Null(result.exception);
            ParseTreeHelper.TreesAreEqual(expectedTree, result.ToStringTree(parser));
        }
    }
}
