namespace AutoCAD_BoardSorter.Models
{
    internal sealed class BoardCoatingSlots
    {
        public string P1 { get; set; }
        public string P2 { get; set; }
        public string L1 { get; set; }
        public string L2 { get; set; }
        public string W1 { get; set; }
        public string W2 { get; set; }

        public bool HasAny
        {
            get
            {
                return !string.IsNullOrWhiteSpace(P1)
                    || !string.IsNullOrWhiteSpace(P2)
                    || !string.IsNullOrWhiteSpace(L1)
                    || !string.IsNullOrWhiteSpace(L2)
                    || !string.IsNullOrWhiteSpace(W1)
                    || !string.IsNullOrWhiteSpace(W2);
            }
        }
    }
}
