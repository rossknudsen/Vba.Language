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

        public bool IsActiveRegion { get { throw new NotImplementedException(); } }
    }

    internal static class ConditionalExtensions
    {
        internal static ConditionalBlock CurrentBlock(this IList<ConditionalBlock> list)
        {
            var current = list.LastOrDefault();
            if (current == null || current.IsComplete)
            {
                return null;
            }
            return current.CurrentBlock();
        }

        internal static ConditionalBlock CurrentBlock(this ConditionalBlock block)
        {
            if (block.IsComplete)
            {
                return null;
            }
            if (block.Else != null 
                && block.Else.HasIncompleteChildren())
            {
                return ((IList<ConditionalBlock>)block.Else.ChildBlocks).CurrentBlock();
            }
            if (block.ElseIfs.Count > 0)
            {
                var lastElseIf = block.ElseIfs.Last();
                if (lastElseIf.HasIncompleteChildren())
                {
                    return ((IList<ConditionalBlock>)lastElseIf.ChildBlocks).CurrentBlock();
                }
            }
            if (block.If != null
                && block.If.HasIncompleteChildren())
            {
                return ((IList<ConditionalBlock>)block.If.ChildBlocks).CurrentBlock();
            }
            // If we get to here then there are no child blocks and this must be the current one.
            return block;
        }

        internal static bool HasIncompleteChildren<T>(this IConditionalNode<T> node) where T : ParserRuleContext
        {
            var lastChild = node.ChildBlocks.LastOrDefault();
            return lastChild != null && !lastChild.IsComplete;
        }
    }
}