using System;

using Vba.Grammars;

namespace Vba.Language
{
    internal class ExpressionEvaluator : PreprocessorBaseVisitor<object>
    {
        private readonly ConstantsDictionary constants;

        public ExpressionEvaluator(ConstantsDictionary constants)
        {
            this.constants = constants;
        }

        public override object VisitExpression(PreprocessorParser.ExpressionContext context)
        {
            // TODO need to implement this method and all other expression methods
            throw new NotImplementedException();
        }
    }
}