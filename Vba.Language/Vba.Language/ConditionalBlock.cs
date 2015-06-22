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

        internal void AddNode<T>(T statement) where T : ParserRuleContext
        {
            if (statement == null)
            {
                throw new ArgumentNullException("statement");
            }
            dynamic dStat = statement;
            CheckStatementSyntax(dStat);
            AddNode(dStat);
        }

        private void AddNode(IfStatement statement)
        {
            If = new ConditionalNode<IfStatement>(statement);
        }

        private void AddNode(ElseIfStatement statement)
        {
            ElseIfs.Add(new ConditionalNode<ElseIfStatement>(statement));
        }

        private void AddNode(ElseStatement statement)
        {
            Else = new ConditionalNode<ElseStatement>(statement);
        }

        private void AddNode(EndIfStatement statement)
        {
            EndIf = new ConditionalNode<EndIfStatement>(statement);
        }

        private void CheckStatementSyntax<T>(T statement) where T : class
        {
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