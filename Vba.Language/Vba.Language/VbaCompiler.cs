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
        private static Func<VbaParser, ParserRuleContext> defaultRule = p => p.module();

        public VbaParseResult CompileSource(string source)
        {
            return CompileSource(new StringReader(source));
        }

        public VbaParseResult CompileSource(string source, Func<VbaParser, ParserRuleContext> rule)
        {
            return CompileSource(new StringReader(source), rule);
        }

        public VbaParseResult CompileSource(TextReader source)
        {
            return CompileSource(source, defaultRule);
        }

        public VbaParseResult CompileSource(TextReader source, Func<VbaParser, ParserRuleContext> rule)
        {
            var headerReader = new ModuleHeaderTextReader(source);
            var preprocessor = new SourceTextReader(headerReader);
            var input = new AntlrInputStream(preprocessor);
            var lexer = new VbaLexer(input);
            var proxy = new LineAdjustProxy(lexer, () => headerReader.Line);
            var tokens = new CommonTokenStream(proxy);
            var parser = new VbaParser(tokens);

            return new VbaParseResult(rule(parser), parser);
        }
    }
}
