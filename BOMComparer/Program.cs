using BOMComparer;

string filePath = @"C:\Users\Arnas\Documents\GitHub\BOMComparer\BOMComparer\bin\Debug\net6.0\BOM_B.xlsx";
string filePath2 = @"C:\Users\Arnas\Documents\GitHub\BOMComparer\BOMComparer\bin\Debug\net6.0\BOM_A.xls";

ExcelReader exr = new();
ExcelReader exrr = new();
var bomFile = exr.ReadBomFile(filePath);
var bomFile2 = exrr.ReadBomFile(filePath2);

Comparer comparer= new();
var compared = comparer.ComparedBomFile(bomFile, bomFile2);
ExcelWriter excelWriter = new();
excelWriter.WriteExcelFile(compared);