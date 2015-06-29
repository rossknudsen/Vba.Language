using System.Collections.Generic;

using Antlr4.Runtime;

namespace Vba.Language
{
    public interface IConditionalNode<T> where T : ParserRuleContext
    {
        IList<IConditionalBlock> ChildBlocks { get; }

        object Result { get; }
    }
}
