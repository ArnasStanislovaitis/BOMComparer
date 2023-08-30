
namespace BOMComparer.Models
{
    public class BomFile
    {
        public string Name { get; set; } = string.Empty;
        public Dictionary<string, BomFileRow> BomFileRow { get; set; } = new();
    }
}