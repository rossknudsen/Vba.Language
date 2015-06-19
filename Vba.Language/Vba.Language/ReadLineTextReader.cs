using System.IO;

namespace Vba.Language
{
    /// <summary>
    /// TextReader class requires an implementation of both Read() and Peek(). 
    /// This is a helper class that implements both of those based off a derived implementation of ReadLine().
    /// This is useful if a derived TextReader can implement ReadLine() more easily than just Read().
    /// <see cref="http://blogs.msdn.com/b/jmstall/archive/2005/08/06/readlinetextreader.aspx"/>
    /// </summary>
    public abstract class ReadLineTextReader : TextReader
    {
        // The default TextReader.Peek() implementation just returns -1. How lame!
        // We can build a real implementation on top of Read().
        public override int Peek()
        {
            FillCharCache();
            return m_charCache;
        }

        // Reads one character. TextReader() demands this be implemented.
        public override int Read()
        {
            FillCharCache();
            int ch = m_charCache;
            ClearCharCache();
            return ch;
        }

        #region Character cache support
        int m_charCache = -2; // -2 means the cache is empty. -1 means eof.
        void ClearCharCache()
        {
            m_charCache = -2;
        }
        void FillCharCache()
        {
            if (m_charCache != -2) return; // cache is already full
            m_charCache = GetNextCharWorker();
        }
        #endregion

        #region Worker to get next signle character from a ReadLine()-based source
        // The whole point of this helper class is that the derived class is going to 
        // implement ReadLine() instead of Read(). So mark that we don't want to use TextReader's 
        // default implementation of ReadLine(). Null return means eof.
        public abstract override string ReadLine();

        // Gets the next char and advances the cursor.
        int GetNextCharWorker()
        {
            // Return the current character
            if (m_line == null)
            {
                m_line = ReadLine(); // virtual
                m_idx = 0;
                if (m_line == null)
                {
                    return -1; // eof
                }
                m_line += "\r\n"; // need to readd the newline that ReadLine() stripped
            }
            char c = m_line[m_idx];
            m_idx++;
            if (m_idx >= m_line.Length)
            {
                m_line = null; // tell us next time around to get a new line.
            }
            return c;
        }

        // Current buffer
        int m_idx = int.MaxValue;
        string m_line;
        #endregion
    }
}