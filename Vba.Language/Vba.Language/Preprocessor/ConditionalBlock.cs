using System;
using System.Collections.Generic;

using Antlr4.Runtime;

using IfStatement = Vba.Grammars.PreprocessorParser.IfStatementContext;
using ElseIfStatement = Vba.Grammars.PreprocessorParser.ElseIfStatementContext;
using ElseStatement = Vba.Grammars.PreprocessorParser.ElseStatementContext;
using EndIfStatement = Vba.Grammars.PreprocessorParser.EndIfStatementContext;

namespace Vba.Language.Preprocessor
{
    internal class ConditionalBlock : IConditionalBlock
    {
        public ConditionalBlock()
        {
            ElseIfs = new List<IConditionalNode<ElseIfStatement>>();
        }

        public IConditionalNode<IfStatement> If { get; internal set; }

        public IList<IConditionalNode<ElseIfStatement>> ElseIfs { get; private set; }

        public IConditionalNode<ElseStatement> Else { get; internal set; }

        public IConditionalNode<EndIfStatement> EndIf { get; internal set; }

        public bool IsComplete { get { return EndIf != null; } }

        internal void AddNode(IfStatement statement, object result)
        {
            CheckStatementSyntax(statement);
            If = new ConditionalNode<IfStatement>(statement, result);
        }

        internal void AddNode(ElseIfStatement statement, object result)
        {
            CheckStatementSyntax(statement);
            ElseIfs.Add(new ConditionalNode<ElseIfStatement>(statement, result));
        }

        internal void AddNode(ElseStatement statement, object result)
        {
            CheckStatementSyntax(statement);
            Else = new ConditionalNode<ElseStatement>(statement, result);
        }

        internal void AddNode(EndIfStatement statement, object result)
        {
            CheckStatementSyntax(statement);
            EndIf = new ConditionalNode<EndIfStatement>(statement, result);
        }

        private void CheckStatementSyntax<T>(T statement) where T : class
        {
            if (statement == null)
            {
                throw new ArgumentNullException("statement");
            }
            var statementType = statement.GetType();
            if (statementType == typeof(IfStatement))
            {
                if (If != null)
                {
                    // TODO Raise critical syntax error.
                }
                if (Else != null)
                {
                    // TODO raise critical syntax error.
                }
                if (ElseIfs.Count > 0)
                {
                    // TODO raise critical syntax error.
                }
                if (EndIf != null)
                {
                    // TODO raise critical syntax error.
                }
                return;
            }

            // These checks must hold for all other statement types.
            if (If == null)
            {
                // TODO Raise critical syntax error.
            }
            if (EndIf != null)
            {
                // TODO Raise critical syntax error.
            }

            if (statementType == typeof(ElseIfStatement))
            {
                if (Else != null)
                {
                    // TODO raise critical syntax error.
                }
            }
            else if (statementType == typeof(ElseStatement))
            {
                if (Else != null)
                {
                    // TODO raise critical syntax error.
                }
            }
        }
    }
}