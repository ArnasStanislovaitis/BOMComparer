using BOMComparer.Models;
using BOMComparer;

namespace ComparerAPI
{
    public class BomFileService : IBomFileService
    {
        private readonly ExcelReader _excelReader;
        private readonly Comparer _comparer;
        private readonly ExcelWriter _excelWriter;
        public BomFileService(ExcelReader excelReader, Comparer comparer, ExcelWriter excelWriter)
        {
            _excelReader = excelReader;
            _comparer = comparer;
            _excelWriter = excelWriter;
        }

        public ComparedBomFile CompareBomFiles(string sourcePath, string targetPath)
        {
            var sourceBomFile = _excelReader.ReadBomFile(sourcePath);
            var targetBomFile = _excelReader.ReadBomFile(targetPath);
            return _comparer.ComparedBomFile(sourceBomFile, targetBomFile);
        }

        public byte[] WriteComparedBomToExcel(ComparedBomFile comparedBomFile)
        {
            return _excelWriter.WriteExcelFile(comparedBomFile);
        }
    }
}