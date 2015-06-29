using System;
using System.Collections.Generic;
using System.Linq;

using Antlr4.Runtime;

using Vba.Grammars;

namespace Vba.Language.Preprocessor
{
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

        internal static bool IsActive(this ConditionalBlock block)
        {
            if (block.EndIf != null)
            {
                return false;
            }
            if (block.Else != null)
            {
                return ((ConditionalNode<PreprocessorParser.ElseStatementContext>)block.Else).IsActive();
            }
            if (block.ElseIfs.Count > 0)
            {
                return ((ConditionalNode<PreprocessorParser.ElseIfStatementContext>)block.ElseIfs.Last()).IsActive();
            }
            if (block.If != null)
            {
                return ((ConditionalNode<PreprocessorParser.IfStatementContext>)block.If).IsActive();
            }
            throw new Exception("Unknown node type.");
        }

        internal static bool IsActive<T>(this ConditionalNode<T> node) where T : ParserRuleContext
        {
            if (node.Result == null)
            {
                return false; // check what the correct response is here.
            }
            if (node.Result is bool)
            {
                return (bool)node.Result;
            }
            // TODO need to add additional type checks etc.
            throw new NotImplementedException();
        }

        internal static bool IsElseStatementActive(this ConditionalBlock block)
        {
            if (IsConditionalNodeActive(block.If))
            {
                return false;
            }
            if (block.ElseIfs.Any(s => IsConditionalNodeActive(s)))
            {
                return false;
            }
            return true;
        }

        private static bool IsConditionalNodeActive<T>(IConditionalNode<T> node) where T : ParserRuleContext
        {
            if (node == null || (node.Result == null))
            {
                throw new NotImplementedException("IsConditionalNodeActive");
            }
            if (node.Result is bool)
            {
                return (bool)node.Result;
            }
            // TODO handle other data types.
            throw new NotImplementedException("IsConditionalNodeActive");
        }
    }
}