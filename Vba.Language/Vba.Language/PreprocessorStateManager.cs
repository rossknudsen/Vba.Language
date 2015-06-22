using System;
using System.Collections.Generic;
using System.Linq;

using Antlr4.Runtime;

using Vba.Grammars;

namespace Vba.Language.Preprocessor
{
    internal class PreprocessorStateManager
    {
        private IList<ConditionalBlock> conditionalBlocks; 
        private ConstantsDictionary compilerConstants;
        PreprocessorLexer lexer;
        PreprocessorParser parser; 

        public PreprocessorStateManager()
        {
            conditionalBlocks = new List<ConditionalBlock>();
            compilerConstants = new ConstantsDictionary();
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

            // TODO handle syntax errors.

            // depending on the type of statement take the appropriate action.
            if (tree.constantDeclaration() != null)
            {
                compilerConstants.Add(tree.constantDeclaration());
                // TODO handle syntax errors.
            }
            else if (tree.moduleAttribute() != null)
            {
                // TODO check if we are still parsing the header.
                // TODO add attribute to the collection.
                // TODO handle syntax errors.
            }
            else
            {
                var current = conditionalBlocks.LastOrDefault();
                if (current == null || current.IsComplete)
                {
                    current = new ConditionalBlock();
                    conditionalBlocks.Add(current);
                }
                current.AddNode(tree);  // this may not be correct, may need a subtree.
                // TODO handle syntax errors.
            }
        }

        public bool IsActiveRegion { get; private set; }
    }
}