using Antlr4.Runtime;

using Vba.Grammars;

namespace Vba.Language.Tests.Preprocessor
{
    internal static class PreprocessorHelper
    {
        internal static PreprocessorParser BuildPreprocessorParser(string source)
        {
            var errorHandler = new TestErrorListener();

            var input = new AntlrInputStream(source);

            var lexer = new PreprocessorLexer(input);
            lexer.AddErrorListener(errorHandler);

            var tokens = new CommonTokenStream(lexer);

            var parser = new PreprocessorParser(tokens);
            parser.AddErrorListener(errorHandler);
            parser.AddErrorListener(new DiagnosticErrorListener());
            parser.ErrorHandler = new BailErrorStrategy();
            return parser;
        }
    }
}
