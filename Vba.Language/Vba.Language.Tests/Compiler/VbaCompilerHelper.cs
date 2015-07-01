using Antlr4.Runtime;

using Vba.Grammars;

namespace Vba.Language.Tests.Compiler
{
    public class VbaCompilerHelper
    {
        internal static VbaParser BuildPreprocessorParser(string source)
        {
            var errorHandler = new TestErrorListener();

            var input = new AntlrInputStream(source);

            var lexer = new VbaLexer(input);
            lexer.AddErrorListener(errorHandler);

            var tokens = new CommonTokenStream(lexer);

            var parser = new VbaParser(tokens);
            parser.AddErrorListener(errorHandler);
            parser.AddErrorListener(new DiagnosticErrorListener());
            parser.ErrorHandler = new BailErrorStrategy();
            return parser;
        }
    }
}
