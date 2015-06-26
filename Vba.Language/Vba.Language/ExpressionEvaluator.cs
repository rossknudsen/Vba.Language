using System;
using System.Diagnostics;

using Vba.Grammars;

namespace Vba.Language
{
    internal class ExpressionEvaluator : PreprocessorBaseVisitor<object>
    {
        private readonly ConstantsDictionary constants;
        private StringComparison stringComparison;

        public ExpressionEvaluator(ConstantsDictionary constants)
        {
            this.constants = constants;
        }

        public override object VisitExpression(PreprocessorParser.ExpressionContext context)
        {
            if (context.comparisonOperator() != null
                && context.expression().Count == 2)
            {
                dynamic left = VisitExpression(context.expression(0));   // could be any expression, bool, string...
                dynamic right = VisitExpression(context.expression(1));  // could be any expression, bool, string...
                return CompareValues(left, right, context.comparisonOperator().GetText());  // This could fail if we don't have a matching overload.
            }
            if (context.boolExpression() != null)
            {
                return VisitBoolExpression(context.boolExpression());
            }
            if (context.stringExpression() != null)
            {
                return VisitStringExpression(context.stringExpression());
            }
            // TODO need to implement this method and all other expression methods
            throw new NotImplementedException("VisitExpression");
        }

        #region Boolean Expressions

        public override object VisitBoolExpression(PreprocessorParser.BoolExpressionContext context)
        {
            // check for Not unary operator.
            if (context.Not() != null 
                && context.boolExpression() != null
                && context.boolExpression(0) != null)
            {
                return !(bool)VisitBoolExpression(context.boolExpression(0));
            }

            // check for And binary operator.
            if (context.And() != null 
                && context.boolExpression() != null 
                && context.boolExpression().Count == 2)
            {
                var left = (bool)VisitBoolExpression(context.boolExpression(0));
                var right = (bool)VisitBoolExpression(context.boolExpression(1));
                return left && right;
            }

            // check for Or binary operator.
            if (context.Or() != null
                && context.boolExpression() != null
                && context.boolExpression().Count == 2)
            {
                var left = (bool)VisitBoolExpression(context.boolExpression(0));
                var right = (bool)VisitBoolExpression(context.boolExpression(1));
                return left || right;
            }

            // check for Xor binary operator.
            if (context.Xor() != null
                && context.boolExpression() != null
                && context.boolExpression().Count == 2)
            {
                var left = (bool)VisitBoolExpression(context.boolExpression(0));
                var right = (bool)VisitBoolExpression(context.boolExpression(1));
                return left ^ right;
            }

            // check for boolean literal.
            if (context.boolLiteral() != null)
            {
                return VisitBoolLiteral(context.boolLiteral());
            }
            throw new NotImplementedException("VisitBoolExpression");
        }

        public override object VisitBoolLiteral(PreprocessorParser.BoolLiteralContext context)
        {
            if (context.True() != null)
            {
                return true;
            }
            if (context.False() != null)
            {
                return false;
            }
            throw new Exception("VisitBoolLiteral");
        }

        #endregion

        #region String Expressions

        public override object VisitStringExpression(PreprocessorParser.StringExpressionContext context)
        {
            if ((context.AMP() != null || context.PLUS() != null) 
                && context.stringExpression().Count == 2)
            {
                var left = (string)VisitStringExpression(context.stringExpression(0));
                var right = (string)VisitStringExpression(context.stringExpression(1));
                return left + right;
            }
            if (context.Like() != null)
            {
                throw new NotImplementedException("Need to implement like comparison");
            }
            if (context.StringLiteral() != null)
            {
                var text = context.StringLiteral().GetText();
                Debug.Assert(text.Length > 2);
                return text.Substring(1, text.Length - 2);
            }
            throw new NotImplementedException("VisitStringExpression: " + context.GetText());
        }

        #endregion

        #region Comparison Expressions

        private object CompareValues(string left, string right, string op)
        {
            var areEqual = AreStringsEqual(left, right);
            var leftIsLessThan = IsLeftStringLessThan(left, right);

            switch (op)
            {
                case "=":
                    return areEqual;
                case "<":
                    return !areEqual && leftIsLessThan;
                case ">":
                    return !areEqual && !leftIsLessThan;
                case "<=":
                    return areEqual || leftIsLessThan;
                case ">=":
                    return areEqual || !leftIsLessThan;
                case "><":
                case "<>":
                    return !areEqual;
                default:
                    throw new NotImplementedException();
            }
        }

        private bool AreStringsEqual(string left, string right)
        {
            return String.Equals(left, right, stringComparison);
        }

        private bool IsLeftStringLessThan(string left, string right)
        {
            // TODO this may not be correct.  Also need to check against VBA implementation vs .Net Framework implementation.
            return string.Compare(left, right, stringComparison) < 0;
        }

        private object CompareValues(bool left, bool right, string op)
        {
            switch (op)
            {
                case "=":
                    return left == right;
                case "<":
                    return ConvertBoolToInt(left) < ConvertBoolToInt(right);
                case ">":
                    return ConvertBoolToInt(left) > ConvertBoolToInt(right);
                case "<=":
                    return ConvertBoolToInt(left) <= ConvertBoolToInt(right);
                case ">=":
                    return ConvertBoolToInt(left) >= ConvertBoolToInt(right);
                case "><":
                case "<>":
                    return left != right;
                default:
                    throw new NotImplementedException();
            }
        }

        private int ConvertBoolToInt(bool value)
        {
            // VBAL p148: True is considered less than False
            if (value)
            {
                return -1;  // true
            }
            return 0;       // false
        }
        #endregion
    }
}