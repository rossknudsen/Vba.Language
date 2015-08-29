using System.Windows.Input;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using Vba.Language;

namespace GrammarTester.ViewModel
{
    /// <summary>
    /// This class contains properties that the main View can data bind to.
    /// <para>
    /// Use the <strong>mvvminpc</strong> snippet to add bindable properties to this ViewModel.
    /// </para>
    /// <para>
    /// You can also use Blend to data bind with the tool's support.
    /// </para>
    /// <para>
    /// See http://www.galasoft.ch/mvvm
    /// </para>
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        private string _source = "";
        private string _output = "";

        /// <summary>
        /// Initializes a new instance of the MainViewModel class.
        /// </summary>
        public MainViewModel()
        {
            ////if (IsInDesignMode)
            ////{
            ////    // Code runs in Blend --> create design time data.
            ////}
            ////else
            ////{
            ////    // Code runs "for real"
            ////}
            
            ParseCommand = new RelayCommand(ParseInput);
        }

        public string Source
        {
            get { return _source; }
            set
            {
                _source = value;
                RaisePropertyChanged();
            }
        }

        public string Output
        {
            get { return _output; }
            private set
            {
                _output = value;
                RaisePropertyChanged();
            }
        }

        public ICommand ParseCommand { get; private set; } 

        private void ParseInput()
        {
            var compiler = new VbaCompiler();
            var result = compiler.CompileSource(Source);
            Output = result.ToString();
        }
    }
}