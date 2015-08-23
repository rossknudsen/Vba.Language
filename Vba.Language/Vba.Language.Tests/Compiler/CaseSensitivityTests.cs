using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Vba.Grammars;
using Xunit;

namespace Vba.Language.Tests.Compiler
{
    public class CaseSensitivityTests
    {
        [Fact]
        public void CanLexKeywordsInsensitively()
        {
            // loop through each keyword and test the lexer's ability to match the case.
            foreach (var keyword in VbaLexer.AllKeywords)
            {
                foreach (var k in GetKeywordCasings(keyword))
                {
                    var lexer = VbaCompilerHelper.BuildVbaLexer(k);

                    var token = lexer.NextToken();
                    Assert.Equal(VbaLexer.ConvertTokenNameToValue(keyword), token.Type);
                }
            }
        }

        private static IEnumerable<string> GetKeywordCasings(string keyword)
        {
            yield return keyword;
            yield return keyword.ToLower();
            yield return keyword.ToUpper();
            // TODO consider whether any other alternatives are needed.
        }
    }
}
