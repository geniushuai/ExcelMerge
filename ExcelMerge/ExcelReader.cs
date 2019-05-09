using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    internal class ExcelReader
    {
        internal static IEnumerable<ExcelRow> Read(ISheet sheet)
        {
            var actualRowIndex = 0;
            var maxColCount = 0;         
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                    continue;

                var cells = new List<ExcelCell>();
                int contiuneBlankCount = 0;
                int columnIndex = 0;
                for (; columnIndex < row.LastCellNum; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex);
                    var stringValue = ExcelUtility.GetCellStringValue(cell);                    
                    cells.Add(new ExcelCell(stringValue, columnIndex, rowIndex));
                    if(stringValue=="")
                    {
                        contiuneBlankCount++;
                    }
                    else
                    {
                        contiuneBlankCount = 0;
                    }
                    if(contiuneBlankCount>ExcelHelperCom.ContinueColBlankUpperLimit && columnIndex> maxColCount)
                    {
                        break;
                    }
                }
                if (columnIndex > maxColCount)
                    maxColCount = columnIndex;                

                yield return new ExcelRow(actualRowIndex++, cells);
            }
        }
    }
}
