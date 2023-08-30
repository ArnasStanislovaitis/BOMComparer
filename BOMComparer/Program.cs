using BOMComparer;

string filePath2 = @"C:\Users\Arnas\Desktop\04_2023-08-09\BOM_B.xlsx";
string filePath = @"C:\Users\Arnas\Desktop\04_2023-08-09\BOM_A.xls";

ExcelReader exr = new();
ExcelReader exrr = new();
var bomFile = exr.ReadBomFile(filePath);
var bomFile2 = exrr.ReadBomFile(filePath2);

Comparer comparer= new();
var compared = comparer.ComparedBomFile(bomFile, bomFile2);
ExcelWriter excelWriter = new();
excelWriter.WriteExcelFile(compared);