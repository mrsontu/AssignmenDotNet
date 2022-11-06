namespace WindowsFormsApp1.Entities
{
    public class ExHeader
    {
        public int RowIndex { get; set; }
        public string Value { get; set; }
        public int StartCol { get; set; }
        public int EndCol { get; set; }
    }

    public class ExCell
    {
        public string Value { get; set; }
        public int Col { get; set; }
    }

    public class ExPosition
    {
        public int Row { get; set; }
        public int Col { get; set; }
    }

    public class ExRange
    {
        public ExPosition StartPoint { get; set; }
        public ExPosition EndPoint { get; set; }

        public bool InsideRange(ExPosition position)
        {
            return (StartPoint.Row <= position.Row && position.Row <= EndPoint.Row) &&
                   (StartPoint.Col <= position.Col && position.Col <= EndPoint.Col);
        }
    }
}
