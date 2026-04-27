namespace AutoCAD_BoardSorter.Models
{
    internal sealed class AssemblySceneData
    {
        public string SceneId { get; set; }
        public string SceneName { get; set; }

        public bool HasValue
        {
            get
            {
                return !string.IsNullOrWhiteSpace(SceneId)
                    || !string.IsNullOrWhiteSpace(SceneName);
            }
        }
    }
}
