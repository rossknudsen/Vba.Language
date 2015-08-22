using Antlr4.Runtime;
using Vba.Grammars;

namespace Vba.Language
{
    public class VbaParseResult
    {
        private readonly RuleContext _result;
        private readonly VbaParser _parser;

        internal VbaParseResult(RuleContext result, VbaParser parser)
        {
            _result = result;
            _parser = parser;
        }

        public override string ToString()
        {
            return _result.ToStringTree(_parser);
        }
    }
}