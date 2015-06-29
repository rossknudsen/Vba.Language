using System.Collections.Generic;

using Vba.Grammars;

namespace Vba.Language.Preprocessor
{
    public interface IConditionalBlock
    {
        IConditionalNode<PreprocessorParser.IfStatementContext> If { get; }

        IList<IConditionalNode<PreprocessorParser.ElseIfStatementContext>> ElseIfs { get; }

        IConditionalNode<PreprocessorParser.ElseStatementContext> Else { get; }

        IConditionalNode<PreprocessorParser.EndIfStatementContext> EndIf { get; }

        bool IsComplete { get; }
    }
}