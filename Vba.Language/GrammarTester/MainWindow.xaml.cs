using System.Windows;
using Vba.Language;

namespace GrammarTester
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            ParseCode();
        }

        private void ParseCode()
        {
            var compiler = new VbaCompiler();
            var result = compiler.CompileSource(txtVbaSource.Text);
            txtOutput.Text = result.ToString();
        }
    }
}
