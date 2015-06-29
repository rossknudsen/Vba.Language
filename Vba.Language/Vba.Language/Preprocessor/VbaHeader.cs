using System.Collections.Generic;

using Antlr4.Runtime;

using Vba.Grammars;

namespace Vba.Language.Preprocessor
{
    internal class VbaHeader
    {
        private IList<PreprocessorParser.ClassHeaderLineContext> classHeaderLines;
        private IList<PreprocessorParser.ModuleAttributeContext> attributes;
        private PreprocessorParser parser;
        private PreprocessorLexer lexer;
        private HeaderErrorListener errorHandler;
        private readonly List<SyntaxError> syntaxErrors;

        internal VbaHeader()
        {
            classHeaderLines = new List<PreprocessorParser.ClassHeaderLineContext>();
            attributes = new List<PreprocessorParser.ModuleAttributeContext>();
            syntaxErrors = new List<SyntaxError>();

            lexer = new PreprocessorLexer(new AntlrInputStream(""));
            var tokens = new CommonTokenStream(lexer);
            parser = new PreprocessorParser(tokens);

            errorHandler = new HeaderErrorListener();
            lexer.AddErrorListener(errorHandler);
            parser.AddErrorListener(errorHandler);
        }

        /// <summary>
        /// Receives the current header line as an argument and processes it.  If processing was successful the
        /// method returns true.  
        /// 
        /// If the line is not a header line then the method returns false.
        /// </summary>
        /// <param name="line">The current header line.</param>
        /// <returns>A boolean value whether the processing was successful or not.</returns>
        internal bool ProcessHeaderStatement(string line)
        {
            errorHandler.Reset();
            lexer.SetInputStream(new AntlrInputStream(line));
            var tree = parser.headerLine();

            if (errorHandler.SyntaxErrors.Count > 0)
            {
                syntaxErrors.AddRange(errorHandler.SyntaxErrors);

                // TODO Some syntax errors may be acceptable rather than just returning false here.
                return false;
            }

            if (tree.classHeaderLine() != null)
            {
                // TODO we must have the class header lines in a valid order.
                classHeaderLines.Add(tree.classHeaderLine());
            }
            if (tree.moduleAttribute() != null)
            {
                attributes.Add(tree.moduleAttribute());
            }
            return true;
        }

        public IReadOnlyList<SyntaxError> SyntaxErrors { get { return syntaxErrors.AsReadOnly(); } }
    }
}
