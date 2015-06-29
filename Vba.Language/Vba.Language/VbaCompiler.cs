using System;
using System.IO;

using Antlr4.Runtime;

using Vba.Grammars;
using Vba.Language.Compiler;
using Vba.Language.Preprocessor;

namespace Vba.Language
{
    public class VbaCompiler
    {
        public object CompileSource(string source)
        {
            return CompileSource(new StringReader(source));
        }

        public object CompileSource(TextReader source)
        {
            var preprocessor = new SourceTextReader(source);
            var input = new AntlrInputStream(preprocessor);
            var lexer = new VbaLexer(input);
            var proxy = new LineAdjustProxy(lexer, () => preprocessor.Line);
            var tokens = new CommonTokenStream(proxy);
            var parser = new VbaParser(tokens);

            // TODO this method should return something.
            throw new NotImplementedException();
        }
    }
}
