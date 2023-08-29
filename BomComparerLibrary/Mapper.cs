using BOMComparer.Models;

namespace BOMComparer
{
    public class Mapper
    {
        public ComparedBomFileRow MapToComparedRow(BomFileRow source, ResultsType resultType, string name)
        {
            var comparedSourceFileRow = new ComparedBomFileRow
            {
                Quantity = source.Quantity,
                PartNumber = source.PartNumber,
                Designator = source.Designator,
                Value = source.Value,
                SMD = source.SMD,
                Description = source.Description,
                Manufacturer = source.Manufacturer,
                ManufacturerPartNr = source.ManufacturerPartNr,
                Distributor = source.Distributor,
                DistributorPartNr = source.DistributorPartNr,
                Result = resultType,
                DataSource = name
            };

            return comparedSourceFileRow;
        }
    }
}