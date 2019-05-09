using System;
using System.Linq;
using System.Data;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Collections.Generic;

namespace ExcelMerge
{
    public class ExcelHelperCom
    {

        public static int ContinueRowBlankUpperLimit = 50;
        public static int ContinueColBlankUpperLimit = 10;
        Dictionary<string, int> rowIndexCache;
        Application myExcel;
        Workbook myWorkBook;
        Worksheet mySheet;

        /// <summary>
        /// 构造函数，不创建Excel工作薄
        /// </summary>
        public ExcelHelperCom()
        {
        }

        /// <summary>
        /// 创建Excel工作薄
        /// </summary>
        public void CreateExcel(Dictionary<string, int> keyCache = null)
        {
            myExcel = new Application();
            myWorkBook = myExcel.Application.Workbooks.Add(true);
            if (keyCache != null)
            {
                rowIndexCache = keyCache;
            }
            else
            { 
                rowIndexCache = new Dictionary<string, int>();
            }
        }

        public bool OpenExcel(string path, Dictionary<string,int> keyCache = null, bool readOnly = false)
        {
            /*      
try
{
    myExcel = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
}
catch (System.Runtime.InteropServices.COMException ex)
{
    myExcel = new Application();
}

string ext = Path.GetExtension(path);
//https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
int saveFomart = 52;//.xlsm
if (ext == ".xlsm")
{
    saveFomart = 52;
}
else if (ext == ".xlsx")
{
    saveFomart = 50;
}
else if (ext == ".csv")
{
    saveFomart = 6;
}
else if (ext == ".xls")
{
    saveFomart = 18;
}
else
{ 
    return false;
}
*/
            myExcel = new Application();
            //myExcel.DisplayAlerts = false;
            myWorkBook = myExcel.Application.Workbooks.Open(path, 0, false);            
            mySheet = myWorkBook.Worksheets.get_Item(1);
            mySheet.Activate();
            if (keyCache != null)
            {
                rowIndexCache = keyCache;
            }
            else
            {
                rowIndexCache = new Dictionary<string, int>();
            }
            return true;
        }
        /// <summary>
        /// 显示Excel
        /// </summary>
        public void ShowExcel()
        {
            myExcel.Visible = true;
        }

        public void Save()
        {
            myWorkBook.Save();
        }

        public void ExitApp()
        {
            myWorkBook.Close(true);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mySheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myWorkBook);
            myExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
        }

        /// <summary>
        /// 将数据写入Excel
        /// </summary>
        /// <param name="data">要写入的二维数组数据</param>
        /// <param name="startRow">Excel中的起始行</param>
        /// <param name="startColumn">Excel中的起始列</param>
        public void WriteData(string[,] data, int startRow, int startColumn)
        {
            if (startRow < 1 || startColumn < 1)
                return;
            int rowNumber = data.GetLength(0);
            int columnNumber = data.GetLength(1);

            for (int i = 0; i < rowNumber; i++)
            {
                for (int j = 0; j < columnNumber; j++)
                {
                    //在Excel中，如果某单元格以单引号“'”开头，表示该单元格为纯文本，因此，我们在每个单元格前面加单引号。 
                    mySheet.Cells[startRow + i, startColumn + j] = "'" + data[i, j];
                }
            }
        }

        /// <summary>
        /// 将数据写入Excel
        /// </summary>
        /// <param name="data">要写入的字符串</param>
        /// <param name="starRow">写入的行</param>
        /// <param name="startColumn">写入的列</param>
        public void WriteData(string data, int row, int column)
        {
            if (row < 1 || column < 1)
                return;
            mySheet.Cells[row, column] = data;
            //设置个颜色表示改过
            mySheet.Cells[row, column].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);            
        }

        /// <summary>
        /// 将数据写入Excel
        /// </summary>
        /// <param name="data">要写入的数据表</param>
        /// <param name="startRow">Excel中的起始行</param>
        /// <param name="startColumn">Excel中的起始列</param>
        public void WriteData(System.Data.DataTable data, int startRow, int startColumn)
        {
            if (startRow < 1 || startColumn < 1)
                return;
            for (int i = 0; i <= data.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= data.Columns.Count - 1; j++)
                {
                    //在Excel中，如果某单元格以单引号“'”开头，表示该单元格为纯文本，因此，我们在每个单元格前面加单引号。 
                    myExcel.Cells[startRow + i, startColumn + j] = "'" + data.Rows[i][j].ToString();
                }
            }            
        }

        /// <summary>
        /// 读取指定单元格数据
        /// </summary>
        /// <param name="row">行序号</param>
        /// <param name="column">列序号</param>
        /// <returns>该格的数据</returns>
        public string ReadData(int row, int column)
        {
            if (row < 1 || column < 1 || row > mySheet.Rows.Count || column>mySheet.Columns.Count)
                return null;
            Range range = mySheet.Cells[row, column];
            return range.Text.ToString();
        }

        public void ActiveCell(int row, int column)
        {
            if (row < 1 || column < 1)
                return;
            try
            {
                mySheet.Cells[row, column].Activate();
            }catch
            { }
        }

        /// <summary>
        /// 重命名工作表
        /// </summary>
        /// <param name="sheetNum">工作表序号，从左到右，从1开始</param>
        /// <param name="newSheetName">新的工作表名</param>
        public void ReNameSheet(int sheetNum, string newSheetName)
        {
            Worksheet worksheet = (Worksheet)myExcel.Worksheets[sheetNum];
            worksheet.Name = newSheetName;
        }

        /// <summary>
        /// 重命名工作表
        /// </summary>
        /// <param name="oldSheetName">原有工作表名</param>
        /// <param name="newSheetName">新的工作表名</param>
        public void ReNameSheet(string oldSheetName, string newSheetName)
        {
            Worksheet worksheet = (Worksheet)myExcel.Worksheets[oldSheetName];
            worksheet.Name = newSheetName;
        }

        /// <summary>
        /// 新建工作表
        /// </summary>
        /// <param name="sheetName">工作表名</param>
        public void CreateWorkSheet(string sheetName)
        {
            Worksheet newWorksheet = (Worksheet)myWorkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            newWorksheet.Name = sheetName;
        }

        /// <summary>
        /// 激活工作表
        /// </summary>
        /// <param name="sheetName">工作表名</param>
        public void ActivateSheet(string sheetName)
        {
            mySheet = (Worksheet)myExcel.Worksheets[sheetName];
            mySheet.Activate();            
        }

        /// <summary>
        /// 激活工作表
        /// </summary>
        /// <param name="sheetNum">工作表序号</param>
        public void ActivateSheet(int sheetNum)
        {
            mySheet = (Worksheet)myExcel.Worksheets[sheetNum];
            mySheet.Activate();
        }

        /// <summary>
        /// 删除一个工作表
        /// </summary>
        /// <param name="SheetName">删除的工作表名</param>
        public void DeleteSheet(int sheetNum)
        {
            int mySheetIndex = mySheet.Index;
            

            ((Worksheet)myWorkBook.Worksheets[sheetNum]).Delete();
            if (mySheetIndex == sheetNum)
            {
                mySheet = myWorkBook.Worksheets.get_Item(1);
                mySheet.Activate();
            }
        }

        /// <summary>
        /// 删除一个工作表
        /// </summary>
        /// <param name="SheetName">删除的工作表序号</param>
        public void DeleteSheet(string sheetName)
        {
            string mySheetName = mySheet.Name;
            ((Worksheet)myWorkBook.Worksheets[sheetName]).Delete();
            if (mySheetName == sheetName)
            {
                mySheet = myWorkBook.Worksheets.get_Item(1);
                mySheet.Activate();
            }
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="startRow">起始行</param>
        /// <param name="startColumn">起始列</param>
        /// <param name="endRow">结束行</param>
        /// <param name="endColumn">结束列</param>
        public void CellsUnite(int startRow, int startColumn, int endRow, int endColumn)
        {

            Range range = myExcel.get_Range(myExcel.Cells[startRow, startColumn], myExcel.Cells[endRow, endColumn]);
            range.MergeCells = true;
        }

        //<status, result> status{1:完全找到,-1:完全没找到,0:最接近的位置}
        public Tuple<int,int> GetRowByHeaderText(string rowHeaderText, int suggestRowIndex)
        {
            rowHeaderText = rowHeaderText.Trim();
            int targetRowIndex = 1;
            string data = "";
            if (rowIndexCache.ContainsKey(rowHeaderText))
            {
                targetRowIndex = rowIndexCache[rowHeaderText];
            }
            else
            {
                try
                {
                    int iRowHeader = Convert.ToInt32(rowHeaderText);
                    int targetRowIndexCandidate = 1;
                    int minDistance = Int32.MaxValue;
                    foreach(var item in rowIndexCache)
                    {
                        try
                        {
                            string header = item.Key;
                            int rowIndex = item.Value;
                            int distance = Math.Abs(Convert.ToInt32(header) - iRowHeader);
                            if (distance < minDistance)
                            {
                                minDistance = distance;
                                targetRowIndexCandidate = rowIndex;
                            }
                        }
                        catch
                        {

                        }
                    }
                    targetRowIndex = targetRowIndexCandidate;
                    try
                    {
                        int iSuggestRowHeader = Convert.ToInt32(ReadData(suggestRowIndex, 1).Trim());
                        if (minDistance > Math.Abs(iRowHeader - iSuggestRowHeader))
                        {
                            targetRowIndex = suggestRowIndex;
                        }
                    }
                    catch
                    {
                    }                                            
                }
                catch
                {
                    targetRowIndex = suggestRowIndex;
                }                
            }
            data = ReadData(targetRowIndex, 1).Trim();
            if (data == rowHeaderText)
            {
                rowIndexCache[rowHeaderText] = targetRowIndex;
                return Tuple.Create(1, targetRowIndex);
            }

            int step = 1;//搜索步进
            int nextBlankCount = 0; //用于判断向后搜索是否有太多空行，是的话就停止向前搜索
            bool preEnd = false;//先前搜索是否完成
            bool nextEnd = false;//向后搜索是否完成
            //部分搜索,com操作太慢了
            while (true && step < 100)
            {
                int[] arr = { targetRowIndex - step, targetRowIndex + step };
                foreach (int i in arr)
                {
                    if (preEnd && nextEnd)
                    {
                        return Tuple.Create(-1, -1) ;
                    }
                    if (i < 1)
                    {
                        preEnd = true;
                        continue;
                    }
                    if (i > mySheet.Rows.Count)
                    {
                        nextEnd = true;
                        continue;
                    }
                    if (preEnd && i < targetRowIndex)
                    {
                        //向前搜索已经结束
                        continue;
                    }
                    if (nextEnd && i > targetRowIndex)
                    {
                        //向后搜索已经结束
                        continue;
                    }
                    data = ReadData(i, 1).Trim();
                    if (data != "")
                    {
                        rowIndexCache[data] = i;
                    }
                    if (data == rowHeaderText)
                    {                        
                        return Tuple.Create(1, i);
                    }                    
                    if (i > targetRowIndex)
                    {
                        if (data == "")
                        {
                            nextBlankCount++;
                        }
                        else
                        {
                            nextBlankCount = 0;
                        }
                        if (nextBlankCount > ContinueRowBlankUpperLimit)
                        {
                            nextEnd = true;
                        }
                    }
                }
                step++;
            }
            return Tuple.Create(0, targetRowIndex);       
        }

        public Tuple<int,int> GetColByHeaderText(string colHeaderText, int suggestColIndex)
        {
            int continuousEmtpyColCount = 0;
            for (int i = 1; i <= mySheet.Columns.Count; ++i)
            {
                string data = ReadData(1, i).Trim();
                if (data == colHeaderText.Trim())
                    return Tuple.Create(1, i);
                if(data == "")
                {
                    continuousEmtpyColCount++;
                    if(continuousEmtpyColCount> ContinueColBlankUpperLimit)
                    {
                        return Tuple.Create(-1, -1);
                    }
                }
                else
                {
                    continuousEmtpyColCount = 0;
                }
            }
            return Tuple.Create(-1, -1);
        }

        public int InsertRow(int row=-1, bool before =true)
        {
            if(row<0)
            {
                row = myExcel.ActiveCell.Row;
            }
            if(!before)
            {
                row += 1;
            }
            Range  line = mySheet.Rows[row];
            line.Insert();
            return line.Row-1;
        }

        public int InsertCol(int col=-1, bool before=true)
        {
            if (col < 0)
            {
                col = myExcel.ActiveCell.Row;
            }
            if (!before)
            {
                col += 1;
            }
            Range column = mySheet.Columns[col];
            column.Insert();
            return column.Column-1;
        }
    }
}
