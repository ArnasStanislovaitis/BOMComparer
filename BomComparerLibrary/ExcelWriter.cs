using BOMComparer.Models;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace BOMComparer
{
    public class ExcelWriter
    {
        readonly Styles styles = new();
        public byte[] WriteExcelFile(ComparedBomFile comparedBomFile)
        {            
            using (var excelPackage = new XSSFWorkbook())
            {
                var worksheet = excelPackage.CreateSheet("Sheet1");
                worksheet.SetAutoFilter(new CellRangeAddress(0, 0, 0, 12));

                for (int row = 0; row < comparedBomFile.ComparedBomFileRows.Count; row++)
                {
                    AddRow(comparedBomFile, row, worksheet, excelPackage);
                }

                styles.FormatHeader(worksheet, excelPackage);
                
                using (var fs = new FileStream("output.xlsx", FileMode.Create))
                {
                    excelPackage.Write(fs);                    
                }
                using (var memoryStream = new MemoryStream())
                {
                    excelPackage.Write(memoryStream, true);
                    return memoryStream.ToArray();
                }
            }
        }
        void SetValues(int row, ComparedBomFile bomFile, IRow excelRow)
        {
            excelRow.CreateCell(0).SetCellValue(bomFile.ComparedBomFileRows[row].Quantity);
            excelRow.CreateCell(1).SetCellValue(bomFile.ComparedBomFileRows[row].PartNumber);
            excelRow.CreateCell(3).SetCellValue(bomFile.ComparedBomFileRows[row].Value);
            excelRow.CreateCell(4).SetCellValue(bomFile.ComparedBomFileRows[row].SMD);
            excelRow.CreateCell(5).SetCellValue(bomFile.ComparedBomFileRows[row].Description);
            excelRow.CreateCell(6).SetCellValue(bomFile.ComparedBomFileRows[row].Manufacturer);
            excelRow.CreateCell(7).SetCellValue(bomFile.ComparedBomFileRows[row].ManufacturerPartNr);
            excelRow.CreateCell(8).SetCellValue(bomFile.ComparedBomFileRows[row].Distributor);
            excelRow.CreateCell(9).SetCellValue(bomFile.ComparedBomFileRows[row].DistributorPartNr);
            excelRow.CreateCell(10).SetCellValue((row == 0) ? nameof(ComparedBomFileRow.Result) : bomFile.ComparedBomFileRows[row].Result.ToString());
            excelRow.CreateCell(11).SetCellValue((row == 0) ? nameof(ComparedBomFileRow.DataSource) : bomFile.ComparedBomFileRows[row].DataSource);
        }
        void AddRow(ComparedBomFile bomFile, int rowIndex, ISheet worksheet, XSSFWorkbook excelPackage)
        {
            var excelRow = worksheet.CreateRow(rowIndex);
            SetValues(rowIndex, bomFile, excelRow);

            for (int column = 0; column < 10; column++)
            {
                var cell = excelRow.GetCell(column);

                if (column == 2)
                {
                    var cellDesignator = excelRow.CreateCell(column);
                    var designator = styles.GetStyledDesignator(bomFile, rowIndex);
                    cellDesignator.SetCellValue(designator);

                    continue;
                }
                else
                {
                    var cellStyle = styles.GetCellStyle(bomFile, worksheet, rowIndex, excelPackage, cell);
                    cell.CellStyle = cellStyle;
                }
            }
        }        
     }
}