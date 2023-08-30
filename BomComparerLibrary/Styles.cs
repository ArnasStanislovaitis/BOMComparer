using BOMComparer.Models;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace BOMComparer
{
    public class Styles
    {
        public ICellStyle GetCellStyle(ComparedBomFile bomFile, ISheet worksheet, int row, XSSFWorkbook excelPackage, ICell cell)
        {
            switch (bomFile.ComparedBomFileRows[row].Result)
            {
                case ResultsType.Added:

                    return StyleAdded(excelPackage);

                case ResultsType.Removed:

                    return StyleRemoved(excelPackage);

                case ResultsType.Modified:

                    return ModifiedCellStyle(row, bomFile, cell, excelPackage);

                default:
                    ICellStyle defaultStyle = worksheet.Workbook.CreateCellStyle();

                    return defaultStyle;
            }
        }
        private ICellStyle ModifiedCellStyle(int row, ComparedBomFile bomFile, ICell cell, XSSFWorkbook excelPackage)
        {
            if (bomFile.ComparedBomFileRows[row].ChangedValues.Contains(cell.ToString()!))
            {
                if (row + 1 < bomFile.ComparedBomFileRows.Count &&
                    bomFile.ComparedBomFileRows[row].PartNumber == bomFile.ComparedBomFileRows[row + 1].PartNumber)
                {
                    return StyleRemoved(excelPackage);
                }
                else
                {
                    return StyleAdded(excelPackage);
                }
            }

            return default!;

        }

        XSSFFont SetDesignatorFont(ComparedBomFile bomFile, int row)
        {
            if (row + 1 < bomFile.ComparedBomFileRows.Count && bomFile.ComparedBomFileRows[row].PartNumber == bomFile.ComparedBomFileRows[row + 1].PartNumber
                    || bomFile.ComparedBomFileRows[row].Result == ResultsType.Removed)
            {
                var font1 = new XSSFFont
                {
                    Color = HSSFColor.Red.Index,
                    IsStrikeout = true
                };

                return font1;
            }
            else
            {
                var font2 = new XSSFFont
                {
                    Color = HSSFColor.Green.Index
                };

                return font2;
            }
        }
        public XSSFRichTextString GetStyledDesignator(ComparedBomFile bomFile, int row)
        {
            var richText = new XSSFRichTextString();
            var font = SetDesignatorFont(bomFile, row);

            foreach (var item in bomFile.ComparedBomFileRows[row].Designator)
            {                             

                if (bomFile.ComparedBomFileRows[row].Designator.IndexOf(item) > 0)
                {
                    richText.Append(", ");
                }

                if (bomFile.ComparedBomFileRows[row].ChangedValues.Contains(item)
                    || bomFile.ComparedBomFileRows[row].Result == ResultsType.Removed
                    || bomFile.ComparedBomFileRows[row].Result == ResultsType.Added)
                {
                    richText.Append(item,font);
                }                
                else
                {
                    richText.Append(item);
                }
                
            }

            return richText;
        }
        public void FormatHeader(ISheet worksheet, XSSFWorkbook excelPackage)
        {
            for (int col = 0; col < 12; col++)
            {
                worksheet.GetRow(0).GetCell(col).CellStyle = HeaderStyle(excelPackage);
                worksheet.AutoSizeColumn(col);
                var columnWidth = worksheet.GetColumnWidth(col);
                worksheet.SetColumnWidth(col, columnWidth + columnWidth / 10);
            }
        }
        private ICellStyle StyleAdded(XSSFWorkbook excelPackage)
        {
            var styleAdded = excelPackage.CreateCellStyle();  
            var font2 = excelPackage.CreateFont();
            font2.Color = HSSFColor.Green.Index;
            styleAdded.SetFont(font2);

            return styleAdded;
        }
        private ICellStyle StyleRemoved(XSSFWorkbook excelPackage)
        {
            var styleRemoved = excelPackage.CreateCellStyle();
            var fontRemoved = excelPackage.CreateFont();
            fontRemoved.IsStrikeout = true;
            fontRemoved.Color = HSSFColor.Red.Index;
            styleRemoved.SetFont(fontRemoved);

            return styleRemoved;
        }
        private ICellStyle HeaderStyle(XSSFWorkbook excelPackage)
        {
            var style = excelPackage.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
    }
}