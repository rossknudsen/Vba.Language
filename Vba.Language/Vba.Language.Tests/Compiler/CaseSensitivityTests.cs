using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Vba.Grammars;
using Xunit;

namespace Vba.Language.Tests.Compiler
{
    public class CaseSensitivityTests
    {
        private readonly IDictionary<string, int> constants = new Dictionary<string, int>(); 

        public CaseSensitivityTests()
        {
            // TODO consider whether we should move the keyword information to the VbaParser class.
            // http://stackoverflow.com/a/10261848/2301065

            var type = typeof (VbaParser);
            var fieldInfos = type.GetFields(BindingFlags.Public | BindingFlags.Static);
            var constantInfos = fieldInfos.Where(f => f.IsLiteral && !f.IsInitOnly).ToList();
            foreach (var c in constantInfos)
            {
                var name = c.Name;
                var value = (int)c.GetRawConstantValue();
                constants.Add(name, value);
            }
        }

        [Fact]
        public void CanLexKeywordsInsensitively()
        {
            // loop through each keyword and test the lexer's ability to match the case.
            foreach (var keyword in VbaKeywords.AllKeywords)
            {
                foreach (var k in GetKeywordCasings(keyword))
                {
                    var lexer = VbaCompilerHelper.BuildVbaLexer(k);

                    var token = lexer.NextToken();
                    Assert.Equal(constants[keyword], token.Type);
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
