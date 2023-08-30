using BOMComparer.Models;

namespace BOMComparer
{
    public class Comparer
    {
        private readonly ComparedBomFile _comparedBomFile;
        private readonly Mapper _mapper;

        public Comparer()
        {
            _comparedBomFile = new ComparedBomFile(); ;
            _mapper = new Mapper();
        }

        public ComparedBomFile ComparedBomFile(BomFile source, BomFile target)
        {            
            CompareValues(source, target);
            AddedValues(source, target);

            return _comparedBomFile;
        }
        private void CompareValues(BomFile source, BomFile target)
        {     

            foreach (var lineKey in source.BomFileRow.Keys)
            {

                if (target.BomFileRow.ContainsKey(lineKey) && EqualRows(source.BomFileRow[lineKey], target.BomFileRow[lineKey]))
                {
                    var comparedRow = _mapper.MapToComparedRow(source.BomFileRow[lineKey], ResultsType.Unchanged,target.Name);
                    _comparedBomFile.ComparedBomFileRows.Add(comparedRow);
                }
                else if (target.BomFileRow.ContainsKey(lineKey) && !EqualRows(source.BomFileRow[lineKey], target.BomFileRow[lineKey]))
                {
                    var bim = ModifiedRowComparer(source.BomFileRow[lineKey], target.BomFileRow[lineKey], source.Name,target.Name);
                    _comparedBomFile.ComparedBomFileRows.Add(bim.Item1);
                    _comparedBomFile.ComparedBomFileRows.Add(bim.Item2);                    
                }
                else if (!target.BomFileRow.ContainsKey(lineKey))
                {
                    var comparedRow = _mapper.MapToComparedRow(source.BomFileRow[lineKey], ResultsType.Removed, source.Name);
                    _comparedBomFile.ComparedBomFileRows.Add(comparedRow);                    
                }

            }
        }
        private void AddedValues(BomFile source, BomFile target)
        {
            foreach (var lineKey in target.BomFileRow.Keys)
            {
                if (!source.BomFileRow.ContainsKey(lineKey))
                {
                    var comparedRow = _mapper.MapToComparedRow(target.BomFileRow[lineKey], ResultsType.Added, target.Name);
                    _comparedBomFile.ComparedBomFileRows.Add(comparedRow);                    
                }
            }
        }

        private (ComparedBomFileRow, ComparedBomFileRow) ModifiedRowComparer(BomFileRow sourceFileRow, BomFileRow targetFileRow, string sourceName, string targetName)
        {
            var comparedSourceFileRow = _mapper.MapToComparedRow(sourceFileRow, ResultsType.Modified, sourceName);
            var comparedTargetFileRow = _mapper.MapToComparedRow(targetFileRow, ResultsType.Modified, targetName);

            var propertiesToCompare = typeof(BomFileRow).GetProperties();

            foreach (var property in propertiesToCompare)
            {
                if (property.PropertyType == typeof(List<string>))
                {
                    var designatorDifferences = DesignatorDifferences(sourceFileRow.Designator, targetFileRow.Designator);
                    comparedSourceFileRow.ChangedValues.AddRange(designatorDifferences.Item1);
                    comparedTargetFileRow.ChangedValues.AddRange(designatorDifferences.Item2);

                    continue;
                }
                var sourceValue = property.GetValue(sourceFileRow)?.ToString();
                var targetValue = property.GetValue(targetFileRow)?.ToString();

                if (sourceValue != targetValue)
                {                    
                    comparedSourceFileRow.ChangedValues.Add(sourceValue!);                    
                    comparedTargetFileRow.ChangedValues.Add(targetValue!);
                }
            }

            return (comparedSourceFileRow, comparedTargetFileRow);
        }

        private (List<string>, List<string>) DesignatorDifferences(List<string> source, List<string> target)
        {
            var removedValues = source.Except(target).ToList();
            var addedValues = target.Except(source).ToList();

            return (removedValues, addedValues);
        }


        private bool EqualRows(BomFileRow sourceFileRow, BomFileRow targetFileRow)
        {
            if (sourceFileRow.Quantity == targetFileRow.Quantity &&
               sourceFileRow.PartNumber == targetFileRow.PartNumber &&
               EqualArr(sourceFileRow.Designator, targetFileRow.Designator) &&
               sourceFileRow.Value == targetFileRow.Value &&
               sourceFileRow.SMD == targetFileRow.SMD &&
               sourceFileRow.Description == targetFileRow.Description &&
               sourceFileRow.Manufacturer == targetFileRow.Manufacturer &&
               sourceFileRow.ManufacturerPartNr == targetFileRow.ManufacturerPartNr &&
               sourceFileRow.Distributor == targetFileRow.Distributor &&
               sourceFileRow.DistributorPartNr == targetFileRow.DistributorPartNr
            )
            {
                return true;
            }

            return false;
        }

        private bool EqualArr(List<string> sourceDesignator, List<string> targetDesignator)
        {
            bool areEqual = sourceDesignator.Count == targetDesignator.Count && sourceDesignator.OrderBy(x => x).SequenceEqual(targetDesignator.OrderBy(x => x));

            if (areEqual)
            {
                return true;
            }

            return false;
        }
    }
}