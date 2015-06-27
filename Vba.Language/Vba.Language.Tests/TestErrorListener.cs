using System;

using Antlr4.Runtime;

namespace Vba.Language.Tests
{
    internal class TestErrorListener : IAntlrErrorListener<IToken>, IAntlrErrorListener<int>
    {
        public void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, 
            int charPositionInLine, string msg, RecognitionException e)
        {
            if (e != null)
                throw e;
            throw new Exception(msg);
        }

        public void SyntaxError(IRecognizer recognizer, int offendingSymbol, int line, 
            int charPositionInLine, string msg, RecognitionException e)
        {
            if (e != null)
                throw e;
            throw new Exception(msg);
        }
    }
}
