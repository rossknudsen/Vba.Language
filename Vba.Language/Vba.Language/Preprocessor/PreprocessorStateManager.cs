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
        private ExpressionEvaluator evaluator;
        private PreprocessorLexer lexer;
        private PreprocessorParser parser; 

        public PreprocessorStateManager()
        {
            conditionalBlocks = new List<ConditionalBlock>();
            compilerConstants = new ConstantsDictionary();
            evaluator = new ExpressionEvaluator(compilerConstants);
        }

        internal void ExecutePreprocessorStatement(string line)
        {
            if (line.Contains("\n") || line.Contains("\r"))
            {
                throw new ArgumentException("Parameter line cannot contain line break characters.");
            }

            SetupParser(line);
            var tree = parser.preprocessorStatement();

            // TODO handle syntax errors.

            // depending on the type of statement take the appropriate action.
            if (tree.constantDeclaration() != null)
            {
                compilerConstants.Add(tree.constantDeclaration());
                // TODO handle syntax errors.
            }
            else
            {
                var current = conditionalBlocks.CurrentBlock();
                if (current == null)
                {
                    current = new ConditionalBlock();
                    conditionalBlocks.Add(current);
                }

                if (tree.ifStatement() != null)
                {
                    var result = evaluator.Visit(tree);
                    current.AddNode(tree.ifStatement(), result);
                }
                else if (tree.elseIfStatement() != null)
                {
                    var result = evaluator.Visit(tree);
                    current.AddNode(tree.elseIfStatement(), result);
                }
                else if (tree.elseStatement() != null)
                {
                    var result = current.IsElseStatementActive();
                    current.AddNode(tree.elseStatement(), result);
                }
                else if (tree.endIfStatement() != null)
                {
                    current.AddNode(tree.endIfStatement(), null);
                }
                else
                {
                    throw new Exception("Unknown node type.");
                }
                // TODO handle syntax errors.
            }
        }

        private void SetupParser(string line)
        {
            lexer = new PreprocessorLexer(new AntlrInputStream(line));
            var tokens = new CommonTokenStream(lexer);
            parser = new PreprocessorParser(tokens);
        }

        public bool IsActiveRegion
        {
            get
            {
                var current = conditionalBlocks.CurrentBlock();
                return current == null || current.IsActive();
            }
        }

        /// <summary>
        /// Checks if the given statement is a preprocessor statement.
        /// </summary>
        /// <param name="line">A string representing a line of source code.</param>
        /// <returns>True if the given statement is a preprocessor statement, false otherwise.</returns>
        /// <exception cref="ArgumentException">Thrown if the input contains line break characters.</exception>
        internal static bool IsPreprocessorStatement(string line)
        {
            if (line.Contains("\r") || line.Contains("\n"))
            {
                throw new ArgumentException("Parameter 'line' cannot contain line break characters.");
            }

            var firstChar = (from c in line
                             where c != '\t' && c != ' '
                             select c).FirstOrDefault();

            return firstChar == '#';
        }
    }
}