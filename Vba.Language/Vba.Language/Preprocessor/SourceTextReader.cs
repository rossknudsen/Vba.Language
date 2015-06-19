using System;
using System.IO;
using System.Linq;

namespace Vba.Language.Preprocessor
{
    public class SourceTextReader : ReadLineTextReader
    {
        private TextReader reader;
        private PreprocessorStateManager manager;

        public SourceTextReader(string source)
        {
            manager = new PreprocessorStateManager();
            reader = new StringReader(source);
        }

        public override string ReadLine()
        {
            string line = reader.ReadLine();
            while (line != null) // null indicates we have reached the end of the source.
            {
                if (IsPreprocessorStatement(line))
                {
                    manager.ExecutePreprocessorStatement(line);
                }
                else if (manager.IsActiveRegion)
                {
                    return line;
                }

                line = reader.ReadLine();
            }
        }

        /// <summary>
        /// Checks if the given statement is a preprocessor statement.
        /// </summary>
        /// <param name="line">A string representing a line of source code.</param>
        /// <returns>True if the given statement is a preprocessor statement, false otherwise.</returns>
        /// <exception cref="ArgumentException">Thrown if the input contains line break characters.</exception>
        private bool IsPreprocessorStatement(string line)
        {
            if (line.Contains("\r") || line.Contains("\n"))
            {
                throw new ArgumentException("Parameter 'line' cannot contain line break characters.");
            }

            var firstChar = (from c in line
                             where c != '\t' && c != ' '
                             select c).FirstOrDefault();

            return firstChar == '#';
        }
    }
}
