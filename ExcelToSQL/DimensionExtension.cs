using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelToSQL
{
    public static class DimensionExtension

    {
        //класс для расширение метода по поиску конца строки и столбца
        public static ExcelAddressBase GetValuedDimension(this ExcelWorksheet worksheet)
        {
            var dimension = worksheet.Dimension;
            if (dimension == null) return null;
            var cells = worksheet.Cells[dimension.Address];
            Int32 minRow = 0, minCol = 0, maxRow = 0, maxCol = 0;
            var hasValue = false;
            foreach (var cell in cells.Where(cell => cell.Value != null))
            {
                if (!hasValue)
                {
                    minRow = cell.Start.Row;
                    minCol = cell.Start.Column;
                    maxRow = cell.End.Row;
                    maxCol = cell.End.Column;
                    hasValue = true;
                }
                else
                {
                    if (cell.Start.Column < minCol)
                    {
                        minCol = cell.Start.Column;
                    }
                    if (cell.End.Row > maxRow)
                    {
                        maxRow = cell.End.Row;
                    }
                    if (cell.End.Column > maxCol)
                    {
                        maxCol = cell.End.Column;
                    }
                }
            }
            return hasValue ? new ExcelAddressBase(minRow, minCol, maxRow, maxCol) : null;
        }
    }
}
