using Antlr4.Runtime;
using Vba.Grammars;

namespace Vba.Language
{
    public class VbaParseResult
    {
        private readonly ParserRuleContext _result;
        private readonly VbaParser _parser;

        internal VbaParseResult(ParserRuleContext result, VbaParser parser)
        {
            _result = result;
            _parser = parser;
        }

        public ParserRuleContext ParserRuleContext { get { return _result; } }

        public override string ToString()
        {
            return _result.ToStringTree(_parser);
        }
    }
}