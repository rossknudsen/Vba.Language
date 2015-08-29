using System.Collections.Generic;
using System.Windows.Media;
using Antlr4.Runtime.Tree;

namespace GrammarTester.ViewModel
{
    public class ParseTreeNodeViewModel
    {
        private IParseTree node;

        internal ParseTreeNodeViewModel(IParseTree node)
        {
            this.node = node;
            Name = "not set";
            TextColor = Brushes.Black;

            if (node is IErrorNode)
            {
                TextColor = Brushes.Red;
            }

            var terminalNode = node as ITerminalNode;
            if (terminalNode != null)
            {
                Name = terminalNode.Symbol.Text;
                return;
            }

            Name = node.GetType().Name;
            IRuleNode ruleNode = (IRuleNode)node;
            for (int i = 0; i < ruleNode.ChildCount; i++)
            {
                Children.Add(new ParseTreeNodeViewModel(ruleNode.GetChild(i)));
            }
        }

        public List<ParseTreeNodeViewModel> Children { get; } = new List<ParseTreeNodeViewModel>();

        public string Name { get; }

        public Brush TextColor { get; private set; }
    }
}