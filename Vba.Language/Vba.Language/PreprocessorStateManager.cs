using System;
using System.Collections.Generic;

using Antlr4.Runtime;

using Vba.Grammars;

namespace Vba.Language.Preprocessor
{
    internal class PreprocessorStateManager
    {
        private IList<ConditionalBlock> conditionalBlocks; 
        private ConstantsDictionary compilerConstants;
        private ExpressionEvaluator evaluator;
        PreprocessorLexer lexer;
        PreprocessorParser parser; 

        public PreprocessorStateManager()
        {
            conditionalBlocks = new List<ConditionalBlock>();
            compilerConstants = new ConstantsDictionary();
            evaluator = new ExpressionEvaluator(compilerConstants);
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
                var result = evaluator.Visit(tree);

                var current = conditionalBlocks.CurrentBlock();
                if (current == null)
                {
                    current = new ConditionalBlock();
                    conditionalBlocks.Add(current);
                }

                if (tree.ifStatement() != null)
                {
                    current.AddNode(tree.ifStatement(), result);
                }
                else if (tree.elseIfStatement() != null)
                {
                    current.AddNode(tree.elseIfStatement(), result);
                }
                else if (tree.elseStatement() != null)
                {
                    current.AddNode(tree.elseStatement(), result);
                }
                else if (tree.endIfStatement() != null)
                {
                    current.AddNode(tree.endIfStatement(), result);
                }
                else
                {
                    throw new Exception("Unknown node type.");
                }
                // TODO handle syntax errors.
            }
        }

        public bool IsActiveRegion
        {
            get
            {
                var current = conditionalBlocks.CurrentBlock();
                return current.IsActive();
            }
        }
    }
}