
namespace BOMComparer.Models
{
    public class BomFileRow
    {
        public string? Quantity { get; set; }
        public string? PartNumber { get; set; }
        public List<string> Designator { get; set; } = null!;
        public string Value { get; set; } = null!;
        public string SMD { get; set; } = null!;
        public string Description { get; set; } = null!;
        public string Manufacturer { get; set; } = null!;
        public string ManufacturerPartNr { get; set; } = null!;
        public string Distributor { get; set; } = null!;
        public string DistributorPartNr { get; set; } = null!;
    }
}