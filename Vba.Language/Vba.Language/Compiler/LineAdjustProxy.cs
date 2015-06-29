using System;

using Antlr4.Runtime;

namespace Vba.Language.Compiler
{
    /// <summary>
    /// This class is a proxy for a true ITokenSource and changes the behaviour of the Line method.  It is used
    /// to adjust the line numbers which the VBA lexer sees to match the true line number that the Preprocessor
    /// sees in the source.
    /// </summary>
    internal class LineAdjustProxy : ITokenSource
    {
        private readonly ITokenSource tokens;
        private readonly Func<int> getLineFunc;

        public LineAdjustProxy(ITokenSource tokens, Func<int> getLineFunc)
        {
            this.tokens = tokens;
            this.getLineFunc = getLineFunc;
        }

        public IToken NextToken()
        {
            return tokens.NextToken();
        }

        public int Line
        {
            get { return getLineFunc(); }
        }

        public int Column
        {
            get { return tokens.Column; }
        }

        public ICharStream InputStream
        {
            get { return tokens.InputStream; }
        }

        public string SourceName
        {
            get { return tokens.SourceName; }
        }

        public ITokenFactory TokenFactory
        {
            get { return tokens.TokenFactory; }
            set { tokens.TokenFactory = value; }
        }
    }
}
