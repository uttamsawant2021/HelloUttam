using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using NPOI.HSSF.UserModel;

namespace DDSiSampleforTest
{
    class Program
    {
        static void Main(string[] args)
        {


            string yes1 = "";
            string message = "2020-10-20 11:24:05,688 [85] ERROR WPI.Global [(null)] - Error Caught in Application_Error event Error in: https://wpi.pg.com/WPI/Pages/Default.aspx Error Message:Object reference not set to an instance of an object. Stack Trace:   at WPI.Default.Page_Load(Object sender, EventArgs e)";
            string erMsg = "WPI/Pages/Default.aspx";
            if (!message.Contains(erMsg))
                yes1 = "0";
            else
                yes1 = "1";

            string[] reqSheets = new string[3] { "Measures$", "Triggers$", "Actions$" };
            //string excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelPath + ";Extended Properties=Excel 12.0;Persist Security Info=False";
            DataSet ds = ReadExcelFile("C:\\A557652\\PNG\\IncidentTrackers\\INC3753207 - Some measures stopped to be automatically updated bu DDSi\\IMEA Daily Shipment Report 05.02.2019.xlsx", reqSheets);

        }

        private static DataSet ReadExcelFile(string excelPath, string[] validSheets)
        {
            DataSet dsOutput = new DataSet();
            DataTable dt = null;

            Dictionary<string, int> columnIndices;

            IWorkbook workbook;
            using (FileStream stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(stream);
            }

            foreach (var validSheet in validSheets)
            {
                columnIndices = new Dictionary<string, int>();

                string sheetName = validSheet.TrimEnd('$');
                ISheet sheet = workbook.GetSheet(sheetName);

                if (sheet == null) continue;

                int startRowNumber;
                dt = new DataTable(validSheet);

                if (validSheet == "Actions$")
                {
                    startRowNumber = 0;
                }
                else
                {
                    startRowNumber = 4;

                    for (int i = 0; i < 3; i++)
                    {
                        DataRow dataRow = dt.NewRow();
                        //dataRow.ItemArray = new string[columnIndices.Count]; //columnIndices.Count is "0" at this point... also, the row needs to be ignored.
                        dt.Rows.Add(dataRow);
                    }
                }

                IRow headerRow = sheet.GetRow(startRowNumber);

                foreach (ICell headerCell in headerRow)
                {
                    columnIndices.Add(headerCell.ToString(), headerCell.ColumnIndex);
                    dt.Columns.Add(headerCell.ToString());
                }
                int lastColumn = headerRow.LastCellNum; // The number of columns in the Header Row is always reliable.

                int rowIndex = 0;
                foreach (IRow row in sheet)
                {
                    if (rowIndex++ < startRowNumber) continue;
                    DataRow dataRow = dt.NewRow();

                    // The NPOI library ignores all cells that have been left untouched... (i.e.) cells without any data OR formatting. 
                    // This is by design... Excel internally stores data in the same format.
                    //var cellsArray = new string[columnIndices.Count];
                    //foreach (var cell in row.Cells)
                    //{
                    //    cellsArray[cell.ColumnIndex] = GetFormattedCellValue(cell);
                    //}

                    // If we wish to make sure that all cells (empty or not) are available in the output we need to iterate over each cell:
                    // http://poi.apache.org/spreadsheet/quick-guide.html#Iterator

                    //IFormulaEvaluator eval;
                    //if (workbook is XSSFWorkbook)
                    //    eval = new XSSFFormulaEvaluator(workbook);
                    //else
                    //    eval = new HSSFFormulaEvaluator(workbook);

                    IList<string> lstCells = new List<string>();
                    for (int cn = 0; cn < lastColumn; cn++)
                    {
                        ICell c = row.GetCell(cn, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (c == null)
                        {
                            //lstCells.Add(string.Empty);
                            lstCells.Add(null);
                        }
                        else
                        {
                            lstCells.Add(GetFormattedCellValue(c)); // , eval
                        }
                    }

                    dataRow.ItemArray = lstCells.ToArray();
                    dt.Rows.Add(dataRow);
                }
                if (validSheet == "Actions$")
                    dt.Rows.Remove(dt.Rows[0]);
                dsOutput.Tables.Add(dt);
            }

            return dsOutput;
        }

        private static string GetFormattedCellValue(ICell cell, IFormulaEvaluator eval = null)
        {
            
            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case NPOI.SS.UserModel.CellType.String:
                        return cell.StringCellValue;

                    case NPOI.SS.UserModel.CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            DateTime date = cell.DateCellValue;
                            //return date.ToString("s");
                            return date.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            return cell.NumericCellValue.ToString();
                        }

                    case NPOI.SS.UserModel.CellType.Boolean:
                        return cell.BooleanCellValue ? "TRUE" : "FALSE";

                    case NPOI.SS.UserModel.CellType.Formula:
                       
                        if (eval != null)
                        {
                            if (cell.CachedFormulaResultType == NPOI.SS.UserModel.CellType.String)
                            {
                                return cell.StringCellValue;
                            }
                            if (cell.CachedFormulaResultType == NPOI.SS.UserModel.CellType.Numeric)
                            {
                                if (DateUtil.IsCellDateFormatted(cell))
                                {
                                    return GetFormattedCellValue(eval.EvaluateInCell(cell));
                                }
                                else
                                {
                                    return cell.NumericCellValue.ToString();
                                }
                            }                            
                        }
                        return cell.CellFormula;

                    case NPOI.SS.UserModel.CellType.Error:
                        return FormulaError.ForInt(cell.ErrorCellValue).String;
                }
            }
            return string.Empty;
        }
    }
}
