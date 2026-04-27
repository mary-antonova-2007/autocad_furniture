namespace AutoCAD_BoardSorter.Models
{
    internal sealed class AssemblyToolData
    {
        public string ToolId { get; set; }
        public string ToolName { get; set; }
        public string ToolCode { get; set; }

        public bool HasValue
        {
            get
            {
                return !string.IsNullOrWhiteSpace(ToolId)
                    || !string.IsNullOrWhiteSpace(ToolName)
                    || !string.IsNullOrWhiteSpace(ToolCode);
            }
        }
    }
}
