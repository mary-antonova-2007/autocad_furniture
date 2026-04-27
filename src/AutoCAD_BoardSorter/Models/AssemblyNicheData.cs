namespace AutoCAD_BoardSorter.Models
{
    internal sealed class AssemblyNicheData
    {
        public string NicheId { get; set; }
        public string NicheName { get; set; }
        public string NicheCode { get; set; }

        public bool HasValue
        {
            get
            {
                return !string.IsNullOrWhiteSpace(NicheId)
                    || !string.IsNullOrWhiteSpace(NicheName)
                    || !string.IsNullOrWhiteSpace(NicheCode);
            }
        }
    }
}
