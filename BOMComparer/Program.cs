using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

List<string> data = new();
string path = @"C:\Users\iot3\source\repos\Zaidimelis\bin\Debug\net7.0\BOM_B.xlsx";

string CellLetter = "ABCDEFGHJKLIMNO";

int rowCount = 13;
int colCount;


BomFile bomFile = new();
BomFile bomFile2 = new();
string filePath = @"C:\Users\iot3\Documents\GitHub\BOMComparer\BOMComparer\bin\Debug\net6.0\BOM_B.xlsx";
using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
{
    IWorkbook workbook = new XSSFWorkbook(fs); // For XLSX files

    var sourceFileName = Path.GetFileName(filePath);
    

    ISheet sheet = workbook.GetSheetAt(0); // Assuming you want to read from the first sheet

    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
    {
        IRow row = sheet.GetRow(rowIndex);
        var dictionaryKey = row.GetCell(1).StringCellValue;
        bomFile2.BomFileRoww[dictionaryKey] = new BomFileRow
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
}

string filePath2 = @"C:\Users\iot3\Documents\GitHub\BOMComparer\BOMComparer\bin\Debug\net6.0\BOM_A.xls";
using (FileStream fs = new FileStream(filePath2, FileMode.Open, FileAccess.Read))
{
    //IWorkbook workbook = new XSSFWorkbook(fs); // For XLSX files

    IWorkbook workbook = new HSSFWorkbook(fs); // For XLS files (older Excel formats)

    ISheet sheet = workbook.GetSheetAt(0); // Assuming you want to read from the first sheet

    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
    {
        IRow row = sheet.GetRow(rowIndex);
        var dictionaryKey = row.GetCell(1).StringCellValue;
        bomFile.BomFileRoww[dictionaryKey] = new BomFileRow
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
}

foreach (var item in bomFile.FileRows)
{
    Console.WriteLine($"{item.Quantity} {item.PartNumber} {item.Manufacturer} {item.Distributor} {item.Description} {item.Designator[0]}");
}
Console.WriteLine();

foreach (var item in bomFile2.FileRows)
{
    Console.WriteLine($"{item.Quantity} {item.PartNumber} {item.Value} {item.Designator[0]}");
}
Console.WriteLine("testinam dic");
Console.WriteLine(bomFile.BomFileRoww.Count());
foreach (var item in bomFile2.BomFileRoww)
{
    Console.WriteLine($"{item.Value.Quantity} {item.Value.PartNumber} {item.Value.Manufacturer} {item.Value.Distributor} {item.Value.Description} {item.Value.Designator[0]}");
}

ComparedBomFile comparedBomFile = new();

foreach (var lineKey in bomFile.BomFileRoww.Keys)
{

    if (bomFile2.BomFileRoww.ContainsKey(lineKey) && EqualRows(bomFile.BomFileRoww[lineKey], bomFile2.BomFileRoww[lineKey]))
    {
        var comparedRow = MapToComparedRow(bomFile.BomFileRoww[lineKey], ResultsType.Unchanged);
        comparedBomFile.ComparedBomFileRows.Add(comparedRow);
    }
    else if (bomFile2.BomFileRoww.ContainsKey(lineKey) && !EqualRows(bomFile.BomFileRoww[lineKey], bomFile2.BomFileRoww[lineKey]))
    {
        var bim = ModifiedRowComparer2(bomFile.BomFileRoww[lineKey], bomFile2.BomFileRoww[lineKey]);
        comparedBomFile.ComparedBomFileRows.Add(bim.Item1);
        comparedBomFile.ComparedBomFileRows.Add(bim.Item2);
        Console.WriteLine("modified");
    }
    else if (!bomFile2.BomFileRoww.ContainsKey(lineKey))
    {
        var comparedRow = MapToComparedRow(bomFile.BomFileRoww[lineKey], ResultsType.Removed);
        comparedBomFile.ComparedBomFileRows.Add(comparedRow);
        Console.WriteLine("removed");
    }

}
foreach (var lineKey in bomFile2.BomFileRoww.Keys)
{
    if (!bomFile.BomFileRoww.ContainsKey(lineKey))
    {
        var comparedRow = MapToComparedRow(bomFile2.BomFileRoww[lineKey], ResultsType.Added);
        comparedBomFile.ComparedBomFileRows.Add(comparedRow);
        Console.WriteLine("added");
    }
}

foreach (var item in comparedBomFile.ComparedBomFileRows)
{
    Console.WriteLine($"{item.Quantity} {item.PartNumber} {string.Join(",", item.Designator)} {item.ManufacturerPartNr} {item.Manufacturer} {item.Description} {item.Distributor} {item.DistributorPartNr} {item.Result} ");
}


(ComparedBomFileRow, ComparedBomFileRow) ModifiedRowComparer2(BomFileRow sourceFileRow, BomFileRow targetFileRow)
{
    var comparedSourceFileRow = MapToComparedRow(sourceFileRow, ResultsType.Modified);
    var comparedTargetFileRow = MapToComparedRow(targetFileRow , ResultsType.Modified);    

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
            property.SetValue(comparedSourceFileRow, sourceValue);
            comparedSourceFileRow.ChangedValues.Add(sourceValue);
            property.SetValue(comparedTargetFileRow, targetValue);
            comparedTargetFileRow.ChangedValues.Add(targetValue);
        }
    }

    return (comparedSourceFileRow, comparedTargetFileRow);
}

(ComparedBomFileRow, ComparedBomFileRow) ModifiedRowComparer(BomFileRow sourceFileRow, BomFileRow targetFileRow)
{
    var comparedSourceFileRow = MapToComparedRow(sourceFileRow, ResultsType.Modified);
    var comparedTargetFileRow = MapToComparedRow(targetFileRow, ResultsType.Modified);

    if (sourceFileRow.Quantity != targetFileRow.Quantity)
    {
        comparedSourceFileRow.Quantity = sourceFileRow.Quantity;
        comparedTargetFileRow.Quantity = targetFileRow.Quantity;
    }
    if (sourceFileRow.Value != targetFileRow.Value)
    {
        comparedSourceFileRow.Value = sourceFileRow.Value;
        comparedTargetFileRow.Value = targetFileRow.Value;
    }
    if (sourceFileRow.SMD != targetFileRow.SMD)
    {
        comparedSourceFileRow.SMD = sourceFileRow.SMD;
        comparedTargetFileRow.SMD = targetFileRow.SMD;
    }
    if (!EqualArr(sourceFileRow.Designator, targetFileRow.Designator))
    {
        var results = DesignatorDifferences(sourceFileRow.Designator, targetFileRow.Designator);
        comparedSourceFileRow.Designator = results.Item1;
        comparedTargetFileRow.Designator = results.Item2;
    }

    if (sourceFileRow.Description != targetFileRow.Description)
    {
        comparedSourceFileRow.Description = sourceFileRow.Description;
        comparedTargetFileRow.Description = targetFileRow.Description;
    }
    if (sourceFileRow.Manufacturer != targetFileRow.Manufacturer)
    {
        comparedSourceFileRow.Manufacturer = sourceFileRow.Manufacturer;
        comparedTargetFileRow.Manufacturer = targetFileRow.Manufacturer;
    }
    if (sourceFileRow.ManufacturerPartNr != targetFileRow.ManufacturerPartNr)
    {
        comparedSourceFileRow.ManufacturerPartNr = sourceFileRow.ManufacturerPartNr;
        comparedTargetFileRow.ManufacturerPartNr = targetFileRow.ManufacturerPartNr;
    }
    if (sourceFileRow.Distributor != targetFileRow.Distributor)
    {
        comparedSourceFileRow.Distributor = sourceFileRow.Distributor;
        comparedTargetFileRow.Distributor = targetFileRow.Distributor;

    }
    if (sourceFileRow.DistributorPartNr != targetFileRow.DistributorPartNr)
    {
        comparedSourceFileRow.DistributorPartNr = sourceFileRow.DistributorPartNr;
        comparedTargetFileRow.DistributorPartNr = targetFileRow.DistributorPartNr;
    }

    return (comparedSourceFileRow, comparedTargetFileRow);
}


(List<string>, List<string>,List<string>,List<string>) DesignatorDifferences2(List<string> source, List<string> target)
{
    var removedValues = source.Except(target).ToList();    //removed
    var addedValues = target.Except(source).ToList();      //added   

    //source = source.Where(x => !removedValues.Contains(x)).ToArray();
    //target = target.Where(x => !addedValues.Contains(x)).ToArray();


    source.RemoveAll(removedValues.Contains);
    target.RemoveAll(addedValues.Contains);

    foreach (var item in removedValues)
    {
        source.Add(item);
    }
    foreach (var item in addedValues)
    {
        target.Add(item);
    }
    source.Sort();
    target.Sort();

    return (source, target,removedValues,addedValues);
}

(List<string>, List<string>) DesignatorDifferences(List<string> source, List<string> target)
{
    var removedValues = source.Except(target).ToList();
    var addedValues = target.Except(source).ToList();    

    return (removedValues, addedValues);
}

ComparedBomFileRow MapToComparedRow(BomFileRow source, ResultsType resultType)
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
        Result = resultType
    };

    return comparedSourceFileRow;
}

bool EqualRows(BomFileRow sourceFileRow, BomFileRow targetFileRow)
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

bool EqualArr(List<string> sourceDesignator, List<string> targetDesignator)
{
    bool areEqual = sourceDesignator.Count == targetDesignator.Count && sourceDesignator.OrderBy(x => x).SequenceEqual(targetDesignator.OrderBy(x => x));

    if (areEqual)
    {
        return true;
    }

    return false;
}
var r = typeof(BomFileRow).GetProperties();
Console.WriteLine(r.Count());
foreach (var property in r)
{
    Console.WriteLine(property);
}




foreach (var item in comparedBomFile.ComparedBomFileRows)
{
    foreach ( var item2 in item.ChangedValues)
    {
        Console.WriteLine( item2);
    }
    

    
}



void blah(ISheet worksheet)
{
    for (int row = 1; row < comparedBomFile.ComparedBomFileRows.Count() + 1; row++)
    {
        IRow excelRow = worksheet.CreateRow(row - 1);        
        excelRow.CreateCell(0).SetCellValue(comparedBomFile.ComparedBomFileRows[row - 1].Quantity);
        excelRow.CreateCell(1).SetCellValue(comparedBomFile.ComparedBomFileRows[row - 1].PartNumber);
        excelRow.CreateCell(2).SetCellValue(string.Join(", ", comparedBomFile.ComparedBomFileRows[row - 1].Designator));
        excelRow.CreateCell(3).SetCellValue(comparedBomFile.ComparedBomFileRows[row - 1].Value);
        excelRow.CreateCell(4).SetCellValue(comparedBomFile.ComparedBomFileRows[row - 1].SMD);
        excelRow.CreateCell(5).SetCellValue(comparedBomFile.ComparedBomFileRows[row - 1].Description);
        excelRow.CreateCell(6).SetCellValue(comparedBomFile.ComparedBomFileRows[row - 1].Manufacturer);
        excelRow.CreateCell(7).SetCellValue(comparedBomFile.ComparedBomFileRows[row - 1].ManufacturerPartNr);
        excelRow.CreateCell(8).SetCellValue(comparedBomFile.ComparedBomFileRows[row - 1].Distributor);
        excelRow.CreateCell(9).SetCellValue(comparedBomFile.ComparedBomFileRows[row - 1].DistributorPartNr);
        excelRow.CreateCell(10).SetCellValue("bym");

    }
}


using (var excelPackage = new XSSFWorkbook())
{
    // Add a new worksheet
    ISheet worksheet = excelPackage.CreateSheet("Sheet1");

    //blah(worksheet);

        


    worksheet.SetAutoFilter(new CellRangeAddress(0, 0, 0, 2));
    // Apply formatting
    ICellStyle style = excelPackage.CreateCellStyle();
    style.FillForegroundColor = IndexedColors.BlueGrey.Index;
    style.FillPattern = FillPattern.SolidForeground;
    style.Alignment = HorizontalAlignment.Center;
    style.VerticalAlignment = VerticalAlignment.Center;


    // Green
    ICellStyle styleGreen = excelPackage.CreateCellStyle(); 
    IFont font = excelPackage.CreateFont();
    font.Boldweight = (short)FontBoldWeight.Bold;
    style.SetFont(font);
    IFont font2 = excelPackage.CreateFont();
    font2.Color = HSSFColor.Green.Index; // Set the font color to red
    styleGreen.SetFont(font2);

    //modified
    ICellStyle styleModified = excelPackage.CreateCellStyle();
    IFont font3 = excelPackage.CreateFont();
    font3.Color = HSSFColor.Orange.Index; // Set the font color to red
    styleModified.SetFont(font3);

    //removed
    ICellStyle styleRemoved = excelPackage.CreateCellStyle();
    IFont fontRemoved = excelPackage.CreateFont();
    fontRemoved.IsStrikeout = true;
    fontRemoved.Color = HSSFColor.Red.Index; // Set the font color to red
    styleRemoved.SetFont(fontRemoved);
    /*
    for (int col = 0; col < 10; col++) // Assuming 10 columns
    {
        worksheet.GetRow(0).GetCell(col).CellStyle = style;
        worksheet.AutoSizeColumn(col);       

        int columnWidth = worksheet.GetColumnWidth(col);
        Console.WriteLine(columnWidth);
        worksheet.SetColumnWidth(col, columnWidth + columnWidth / 10);
    }*/
    /*
    for (int i = 0; i < 13; i++)
    {
        
        for (int c = 0; c < 10; c++)
        {   
            if (comparedBomFile.ComparedBomFileRows[i].Result == ResultsType.Added)
            {               
                worksheet.GetRow(i).GetCell(c).CellStyle = styleGreen;                
            }
            else if(comparedBomFile.ComparedBomFileRows[i].Result == ResultsType.Modified)
            {                
                    
                if(c == 2)
                {                    
                        
                    var richText5 = new XSSFRichTextString();
                    foreach (var item in comparedBomFile.ComparedBomFileRows[i].Designator)
                    {
                        var font6 = new XSSFFont();
                        font6.Color = HSSFColor.Green.Index;


                        if (comparedBomFile.ComparedBomFileRows[i].Designator.IndexOf(item) > 0)
                        {
                            richText5.Append(", ");
                        }
                        if (comparedBomFile.ComparedBomFileRows[i].ChangedValues.Contains(item))
                        {
                            richText5.Append(item, (XSSFFont)font6);
                        }
                        else
                        {
                            richText5.Append(item);
                        }                      

                        
                    }
                    Console.WriteLine(richText5);
                    worksheet.GetRow(i).GetCell(c).SetCellValue(richText5);

                }
                else if (comparedBomFile.ComparedBomFileRows[i].ChangedValues.Contains(worksheet.GetRow(i)?.GetCell(c).ToString()!))
                {
                    worksheet.GetRow(i).GetCell(c).CellStyle = styleModified;
                }
                
            }

        }
    }
    */
    for (int i = 0; i < 13; i++)
    {
        for (int c = 0; c < 10; c++)
        {
            switch (comparedBomFile.ComparedBomFileRows[i].Result)
            {
                case ResultsType.Added:
                    worksheet.GetRow(i).GetCell(c).CellStyle = styleGreen;
                    break;
                case ResultsType.Removed:
                    worksheet.GetRow(i).GetCell(c).CellStyle = styleRemoved;
                    break;
                case ResultsType.Modified:
                    if (c == 2)
                    {
                        var richText5 = new XSSFRichTextString();
                        foreach (var item in comparedBomFile.ComparedBomFileRows[i].Designator)
                        {
                            var font6 = new XSSFFont();
                            if(i + 1 < 13 && comparedBomFile.ComparedBomFileRows[i].PartNumber == comparedBomFile.ComparedBomFileRows[i + 1].PartNumber)
                            {
                                font6.Color = HSSFColor.Red.Index;
                                font6.IsStrikeout = true;
                            }
                            else
                            {
                                font6.Color = HSSFColor.Green.Index;
                            }
                            

                            if (comparedBomFile.ComparedBomFileRows[i].Designator.IndexOf(item) > 0)
                            {
                                richText5.Append(", ");
                            }
                            if (comparedBomFile.ComparedBomFileRows[i].ChangedValues.Contains(item))
                            {
                                richText5.Append(item, (XSSFFont)font6);
                            }
                            else
                            {
                                richText5.Append(item);
                            }
                        }
                        Console.WriteLine(richText5);
                        worksheet.GetRow(i).GetCell(c).SetCellValue(richText5);
                    }
                    else if (comparedBomFile.ComparedBomFileRows[i].ChangedValues.Contains(worksheet.GetRow(i)?.GetCell(c).ToString()!))
                    {
                        worksheet.GetRow(i).GetCell(c).CellStyle = styleModified;
                        if (i + 1 < 13 && comparedBomFile.ComparedBomFileRows[i].PartNumber == comparedBomFile.ComparedBomFileRows[i + 1].PartNumber)
                        {
                            worksheet.GetRow(i).GetCell(c).CellStyle = styleRemoved;
                        }
                        else
                        {
                            worksheet.GetRow(i).GetCell(c).CellStyle = styleGreen;
                        }
                    }
                    break;
            }
        }
    }




    // Save the Excel package to a file
    using (var fs = new FileStream("output2.xlsx", FileMode.Create))
    {
        excelPackage.Write(fs);
    }
}
public class BomFile
{
    public List<BomFileRow> FileRows { get; set; } = new();
    public Dictionary<string, BomFileRow> BomFileRoww { get; set; } = new();
}
public class ComparedBomFile
{
    public List<ComparedBomFileRow> ComparedBomFileRows { get; set; } = new();
}
public class BomFileRow
{
    public string? Quantity { get; set; }
    public string? PartNumber { get; set; }
    public List<string> Designator { get; set; }
    public string Value { get; set; }
    public string SMD { get; set; }
    public string Description { get; set; }
    public string Manufacturer { get; set; }
    public string ManufacturerPartNr { get; set; }
    public string Distributor { get; set; }
    public string DistributorPartNr { get; set; }
}

public class SMDd
{
    public string Smd { get; set; }
    public ResultsType Result { get; set; }
}
public class ComparedBomFileRow : BomFileRow
{
    public List<string> ChangedValues { get; set; } = new();
    public string DataSource { get; set; } = null!;
    public ResultsType Result { get; set; }
}

public enum ResultsType
{
    Unchanged,
    Added,
    Modified,
    Removed
}
