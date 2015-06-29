using System.Collections.Generic;

using Antlr4.Runtime;

namespace Vba.Language.Preprocessor
{
    internal class ConditionalNode<T> : IConditionalNode<T> where T : ParserRuleContext
    {
        public ConditionalNode(T statement, object result)
        {
            ChildBlocks = new List<IConditionalBlock>();
            Statement = statement;
            Result = result;
        }

        public IList<IConditionalBlock> ChildBlocks { get; private set; }

        public object Result { get; private set; }

        internal T Statement { get; private set; }
    }
}