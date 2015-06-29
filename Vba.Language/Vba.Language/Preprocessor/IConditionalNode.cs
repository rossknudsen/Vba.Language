using System.Collections.Generic;

using Antlr4.Runtime;

namespace Vba.Language.Preprocessor
{
    public interface IConditionalNode<T> where T : ParserRuleContext
    {
        IList<IConditionalBlock> ChildBlocks { get; }

        object Result { get; }
    }
}
