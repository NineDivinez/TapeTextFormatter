namespace TapeTextFormatter
{
    internal class TapeData
    {
        internal string name;
        internal string returnDate;
        internal string description;

        public TapeData(string _name, string _returnDate, string _description)
        {
            this.name = _name;
            this.returnDate = _returnDate;
            this.description = _description;
        }

        public override string ToString()
        {
            return this.name;
        }
    }
}
