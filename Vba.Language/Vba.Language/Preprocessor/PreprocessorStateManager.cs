using System.Collections;
using System.Collections.Generic;

namespace Vba.Language.Preprocessor
{
    internal class PreprocessorStateManager
    {
        private Stack<PreprocessorState> stateStack;
        private IDictionary<string, object> compilerConstants;
        // member Antlr lexer/parser...

        public PreprocessorStateManager()
        {
            stateStack = new Stack<PreprocessorState>();
            compilerConstants = new Dictionary<string, object>();
        }

        internal void ExecutePreprocessorStatement(string line)
        {
            // use Antlr parser/lexer to process line.

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