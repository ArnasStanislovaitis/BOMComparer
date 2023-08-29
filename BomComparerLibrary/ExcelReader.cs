using BOMComparer.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace BOMComparer
{
    public class ExcelReader
    {
        public BomFile ReadBomFile(string filePath)
        {
            BomFile bomFile = new();
            IWorkbook workbook;
        
            using (FileStream fs = new(filePath, FileMode.Open, FileAccess.Read))
            {  
                var fileName = Path.GetFileName(filePath);
                bomFile.Name = fileName;
                var extension = Path.GetExtension(fileName);

                if(extension == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else
                {
                    workbook = new HSSFWorkbook(fs);
                }
                var sheet = workbook.GetSheetAt(0);
                
                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    var dictionaryKey = row.GetCell(1).StringCellValue;
                    bomFile.BomFileRow[dictionaryKey] = new BomFileRow
                    {
                        Quantity = row.GetCell(0).ToString(),
                        PartNumber = row.GetCell(1).StringCellValue,
                        Designator = row.GetCell(2).StringCellValue.Split(',').Select(s => s.Trim()).ToList(),
                        Value = row.GetCell(3).StringCellValue,
                        SMD = row.GetCell(4).StringCellValue,
                        Description = row.GetCell(5).StringCellValue,
                        Manufacturer = row.GetCell(6).StringCellValue,
                        ManufacturerPartNr = row.GetCell(7).StringCellValue,
                        Distributor = row.GetCell(8).StringCellValue,
                        DistributorPartNr = row.GetCell(9).StringCellValue,
                    };
                }

                return bomFile;
            }
        }
        
    }
}