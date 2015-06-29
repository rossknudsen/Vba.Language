using System;
using System.Diagnostics;

using Vba.Grammars;

namespace Vba.Language
{
    internal class ExpressionEvaluator : PreprocessorBaseVisitor<object>
    {
        protected readonly ConstantsDictionary constants;
        protected StringComparison stringComparison = StringComparison.InvariantCultureIgnoreCase;

        public ExpressionEvaluator(ConstantsDictionary constants)
        {
            this.constants = constants;
        }

        #region Preprocessor Overrides

        public override object VisitPreprocessorStatement(PreprocessorParser.PreprocessorStatementContext context)
        {
            if (context.ifStatement() != null)
            {
                return VisitIfStatement(context.ifStatement());
            }
            if (context.elseIfStatement() != null)
            {
                return VisitElseIfStatement(context.elseIfStatement());
            }
            if (context.elseStatement() != null)
            {
                return VisitElseStatement(context.elseStatement());
            }
            if (context.endIfStatement() != null)
            {
                return VisitEndIfStatement(context.endIfStatement());
            }
            throw new NotImplementedException("VisitPreprocessorStatement");
        }

        public override object VisitIfStatement(PreprocessorParser.IfStatementContext context)
        {
            return VisitExpression(context.expression());
        }

        public override object VisitElseIfStatement(PreprocessorParser.ElseIfStatementContext context)
        {
            return VisitExpression(context.expression());
        }

        public override object VisitElseStatement(PreprocessorParser.ElseStatementContext context)
        {
            throw new Exception("Cannot evaluate Else Statement");
        }

        public override object VisitEndIfStatement(PreprocessorParser.EndIfStatementContext context)
        {
            // This statement can't really be evaluated.  Not sure what we should return.
            return false;
        }

        #endregion

        public override object VisitExpression(PreprocessorParser.ExpressionContext context)
        {
            if (context.comparisonOperator() != null
                && context.expression().Count == 2)
            {
                dynamic left = VisitExpression(context.expression(0));   // could be any expression, bool, string...
                dynamic right = VisitExpression(context.expression(1));  // could be any expression, bool, string...
                return CompareValues(left, right, context.comparisonOperator().GetText());  // This could fail if we don't have a matching overload.
            }
            if (context.arithmeticExpression() != null)
            {
                return VisitArithmeticExpression(context.arithmeticExpression());
            }
            if (context.boolExpression() != null)
            {
                return VisitBoolExpression(context.boolExpression());
            }
            if (context.stringExpression() != null)
            {
                return VisitStringExpression(context.stringExpression());
            }
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

        #region Arithmetic Expressions

        public override object VisitArithmeticExpression(PreprocessorParser.ArithmeticExpressionContext context)
        {
            if (context.MINUS() != null
                && context.arithmeticExpression().Count == 1)
            {
                return -1 * (int)VisitArithmeticExpression(context.arithmeticExpression(0));
            }
            if (context.IntegerLiteral() != null)
            {
                return ParseInteger(context.IntegerLiteral().GetText());
            }
            if (context.FloatLiteral() != null)
            {
                return ParseFloat(context.FloatLiteral().GetText());
            }

            Debug.Assert(context.arithmeticExpression().Count == 2);  // only operations left are binary.
            
            var left = (int)VisitArithmeticExpression(context.arithmeticExpression(0));
            var right = (int)VisitArithmeticExpression(context.arithmeticExpression(1));
            if (context.CARET() != null)
            {
                return (int)Math.Pow(left, right);
            }
            if (context.FS() != null)
            {
                return left / right;
            }
            if (context.BS() != null)
            {
                return left / right;
            }
            if (context.Mod() != null)
            {
                return left % right;
            }
            if (context.STAR() != null)
            {
                return left * right;
            }
            if (context.PLUS() != null)
            {
                return left + right;
            }
            if (context.MINUS() != null)
            {
                return left - right;
            }
            throw new NotImplementedException("VisitArithmeticExpression");
        }

        private static int ParseInteger(string text)
        {
            // TODO need to handle VBA integer definitions (hex/octal/decimal/suffixes).
            return int.Parse(text);
        }

        private static float ParseFloat(string text)
        {
            // TODO need to handle VBA float definitions (exponents/suffixes).
            return float.Parse(text);
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
            return CompareValues(ConvertBoolToInt(left), ConvertBoolToInt(right), op);
        }

        private object CompareValues(bool left, int right, string op)
        {
            return CompareValues(ConvertBoolToInt(left), right, op);
        }

        private object CompareValues(int left, bool right, string op)
        {
            return CompareValues(left, ConvertBoolToInt(right), op);
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

        private object CompareValues(int left, int right, string op)
        {
            switch (op)
            {
                case "=":
                    return left == right;
                case "<":
                    return left < right;
                case ">":
                    return left > right;
                case "<=":
                    return left <= right;
                case ">=":
                    return left >= right;
                case "><":
                case "<>":
                    return left != right;
                default:
                    throw new NotImplementedException();
            }
        }

        private object CompareValues(string left, bool right, string op)
        {
            // TODO consider creating a specific class to represent errors
            // A Type Mismatch error occurs.
            throw new NotImplementedException();
        }

        private object CompareValues(bool left, string right, string op)
        {
            // TODO consider creating a specific class to represent errors
            // A Type Mismatch error occurs.
            throw new NotImplementedException();
        }

        private object CompareValues(int left, string right, string op)
        {
            // TODO consider creating a specific class to represent errors
            // A Type Mismatch error occurs.
            throw new NotImplementedException();
        }

        private object CompareValues(string left, int right, string op)
        {
            // TODO consider creating a specific class to represent errors
            // A Type Mismatch error occurs.
            throw new NotImplementedException();
        }

        #endregion
    }
}