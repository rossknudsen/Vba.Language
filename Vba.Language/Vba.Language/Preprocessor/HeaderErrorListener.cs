using System.Collections.Generic;

using Antlr4.Runtime;

namespace Vba.Language.Preprocessor
{
    internal class HeaderErrorListener : IAntlrErrorListener<int>, IAntlrErrorListener<IToken>
    {
        private readonly List<SyntaxError> syntaxErrors = new List<SyntaxError>();

        public void SyntaxError(IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            syntaxErrors.Add(new SyntaxError(line, charPositionInLine, msg));
        }

        public void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg,
            RecognitionException e)
        {
            syntaxErrors.Add(new SyntaxError(line, charPositionInLine, msg));
        }

        internal void Reset()
        {
            syntaxErrors.Clear();
        }

        public IReadOnlyList<SyntaxError> SyntaxErrors
        {
            get { return syntaxErrors.AsReadOnly(); }
        }
    }
}