using System;
using System.Collections.Generic;

using Antlr4.Runtime;

using Vba.Grammars;

namespace Vba.Language
{
    internal class VbaHeader
    {
        private IList<PreprocessorParser.ClassHeaderLineContext> classHeaderLines;
        private IList<PreprocessorParser.ModuleAttributeContext> attributes;
        private PreprocessorParser parser;

        internal VbaHeader()
        {
            classHeaderLines = new List<PreprocessorParser.ClassHeaderLineContext>();
            attributes = new List<PreprocessorParser.ModuleAttributeContext>();
            var lexer = new PreprocessorLexer(new AntlrInputStream(""));
            var tokens = new CommonTokenStream(lexer);
            parser = new PreprocessorParser(tokens);
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
            var tree = parser.headerLine();
            if (tree.classHeaderLine() != null)
            {
                classHeaderLines.Add(tree.classHeaderLine());
            }
            else if (tree.moduleAttribute() != null)
            {
                attributes.Add(tree.moduleAttribute());
            }
            // TODO catch lines which do not match valid header lines.
            // TODO deal with any syntax errors in the above lines.
            throw new NotImplementedException();
        }
    }
}
