using System.IO;

namespace Vba.Language.Preprocessor
{
    internal sealed class ModuleHeaderTextReader : ReadLineTextReader
    {
        private TextReader reader;
        private VbaHeader header;
        private bool hasReadHeader = false;

        internal ModuleHeaderTextReader(string source) : this(new StringReader(source)) { }

        internal ModuleHeaderTextReader(TextReader reader)
        {
            this.reader = reader;
            header = new VbaHeader();
        }

        public int Line { get; private set; }

        public override string ReadLine()
        {
            // check if we have read the header yet.
            if (!hasReadHeader)
            {
                var line = NextLine();
                while (line != null && header.ProcessHeaderStatement(line))
                {
                    line = NextLine();
                }
                hasReadHeader = true;
                return line;  // note that we must return here because this is not a header line.
            }
            return NextLine();
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