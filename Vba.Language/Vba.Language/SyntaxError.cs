namespace Vba.Language
{
    internal class SyntaxError
    {
        public SyntaxError(int line, int position, string message)
        {
            Line = line;
            Position = position;
            Message = message;
        }

        public int Line { get; private set; }

        public int Position { get; private set; }

        public string Message { get; private set; }
    }
}