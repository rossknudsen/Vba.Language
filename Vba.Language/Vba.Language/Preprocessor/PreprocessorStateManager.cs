using System;
using System.Collections;
using System.Collections.Generic;

using Antlr4.Runtime;

using Vba.Grammars;

namespace Vba.Language.Preprocessor
{
    internal class PreprocessorStateManager
    {
        private Stack<PreprocessorState> stateStack;
        private IDictionary<string, object> compilerConstants;
        PreprocessorLexer lexer;
        PreprocessorParser parser; 

        public PreprocessorStateManager()
        {
            stateStack = new Stack<PreprocessorState>();
            compilerConstants = new Dictionary<string, object>();
            lexer = new PreprocessorLexer(new AntlrInputStream(""));
            var tokens = new CommonTokenStream(lexer);
            parser = new PreprocessorParser(tokens);
        }

        internal void ExecutePreprocessorStatement(string line)
        {
            if (line.Contains("\n") || line.Contains("\r"))
            {
                throw new ArgumentException("Parameter line cannot contain line break characters.");
            }

            var input = new AntlrInputStream(line);
            lexer.SetInputStream(input);
            var tree = parser.preprocessorStatement();

            // TODO how to handle syntax errors.

            // depending on the type of statement take the appropriate action.
        }

        public bool IsActiveRegion { get; private set; }

        private class PreprocessorState
        {
            private object ifStatement;
            private ArrayList elseIfStatements;
            private object elseStatement;
        }
    }
}