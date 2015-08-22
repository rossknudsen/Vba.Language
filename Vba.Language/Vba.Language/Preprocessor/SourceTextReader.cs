using System.IO;

namespace Vba.Language.Preprocessor
{
    internal sealed class SourceTextReader : ReadLineTextReader
    {
        private TextReader reader;
        private PreprocessorStateManager manager;

        public SourceTextReader(string source) : this(new StringReader(source)) { }

        public SourceTextReader(TextReader reader)
        {
            this.reader = reader;
            manager = new PreprocessorStateManager();
            Line = 0;
        }

        public int Line { get; private set; }

        public override string ReadLine()
        {
            // when ReadLine returns null we are at the end of the file.
            var line = NextLine();

            // after we exit the top loop, line will be the first non-header line.
            while (line != null) 
            {
                if (PreprocessorStateManager.IsPreprocessorStatement(line))
                {
                    manager.ExecutePreprocessorStatement(line);
                }
                else if (manager.IsActiveRegion)
                {
                    return line;
                }

                line = NextLine();
            }
            return null;
        }

        private string NextLine()
        {
            var line = reader.ReadLine();
            Line++;
            return line;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (reader != null)
                {
                    reader.Dispose();
                    reader = null;
                }
            }
            base.Dispose(disposing);
        }
    }
}
