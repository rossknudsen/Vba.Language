using System.Collections.Generic;

using Antlr4.Runtime;

namespace Vba.Language.Preprocessor
{
    internal class ConditionalNode<T> : IConditionalNode<T> where T : ParserRuleContext
    {
        public ConditionalNode(T statement)
        {
            ChildBlocks = new List<IConditionalBlock>();
            Statement = statement;
        }

        public IList<IConditionalBlock> ChildBlocks { get; private set; }

        internal T Statement { get; private set; }
    }
}