namespace TapeTextFormatter
{
    internal class TapeData
    {
        internal string tapeName;
        internal string tapeReturnDate;
        internal string tapeDescription;

        public TapeData(string tapeNames, string tapeReturnDates, string tapeDescription)
        {
            this.tapeName = tapeNames;
            this.tapeReturnDate = tapeReturnDates;
            this.tapeDescription = tapeDescription;
        }

        public override string ToString()
        {
            return this.tapeName;
        }
    }
}
