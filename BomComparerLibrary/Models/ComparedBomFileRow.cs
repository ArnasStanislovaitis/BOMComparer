
namespace BOMComparer.Models
{
    public class ComparedBomFileRow
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
        public List<string> ChangedValues { get; set; } = new();
        public string DataSource { get; set; } = null!;
        public ResultsType Result { get; set; } = ResultsType.Unchanged;
    }
}