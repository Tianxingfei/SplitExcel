using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
 
namespace  HelperExtend
{
    public class AsFunctionExcelIO
    {
        /// <summary>
        /// Excel导入
        /// </summary>
        /// <param name="strFileName">excel文档路径</param>
        /// <returns></returns>
        public DataTable ImportExcelToDataTable(string strFileName,int i)
        {
            DataTable dt = new DataTable();
            try
            {
                    IWorkbook workbook;
                    FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read);

                    workbook = WorkbookFactory.Create(file);
                    file.Close();
                    file.Dispose();
                    ISheet sheet = (ISheet)workbook.GetSheetAt(i);
                    dt = ImportDataTable(sheet, 0, true);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return dt;
        }
        /// <summary>
        /// 将制定sheet中的数据导出到datatable中
        /// </summary>
        /// <param name="sheet">需要导出的sheet</param>
        /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
        /// <returns></returns>
        public DataTable ImportDataTable(ISheet sheet, int HeaderRowIndex, bool needHeader)
        {
            DataTable table = new DataTable();
            IRow headerRow;
            int cellCount;
            try
            {
                if (HeaderRowIndex < 0 || !needHeader)
                {
                    headerRow = sheet.GetRow(0);
                    cellCount = headerRow.LastCellNum;

                    for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                    {
                        DataColumn column = new DataColumn(Convert.ToString(i));
                        table.Columns.Add(column);
                    }
                }
                else
                {
                    headerRow = sheet.GetRow(HeaderRowIndex);
                    cellCount = headerRow.LastCellNum;

                    for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                    {
                        if (headerRow.GetCell(i) == null)
                        {
                            if (table.Columns.IndexOf(Convert.ToString(i)) > 0)
                            {
                                DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                                table.Columns.Add(column);
                            }
                            else
                            {
                                DataColumn column = new DataColumn(Convert.ToString(i));
                                table.Columns.Add(column);
                            }

                        }
                        else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0)
                        {
                            DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                            table.Columns.Add(column);
                        }
                        else
                        {
                            DataColumn column = new DataColumn(headerRow.GetCell(i).ToString());
                            table.Columns.Add(column);
                        }
                    }
                }
                int rowCount = sheet.LastRowNum;
                for (int i = (HeaderRowIndex + 1); i <= sheet.LastRowNum; i++)
                {
                    try
                    {
                        IRow row;
                        if (sheet.GetRow(i) == null)
                        {
                            row = sheet.CreateRow(i);
                        }
                        else
                        {
                            row = sheet.GetRow(i);
                        }

                        DataRow dataRow = table.NewRow();

                        for (int j = row.FirstCellNum; j <= cellCount; j++)
                        {
                            try
                            {
                                if (row.GetCell(j) != null)
                                {
                                    switch (row.GetCell(j).CellType)
                                    {
                                        case CellType.STRING:
                                            string str = row.GetCell(j).StringCellValue;
                                            if (str != null && str.Length > 0)
                                            {
                                                dataRow[j] = str.ToString();
                                            }
                                            else
                                            {
                                                dataRow[j] = null;
                                            }
                                            break;
                                        case CellType.NUMERIC:
                                            if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                            {
                                                dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                            }
                                            else
                                            {
                                                dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                            }
                                            break;
                                        case CellType.BOOLEAN:
                                            dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                            break;
                                        case CellType.ERROR:
                                            dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                            break;
                                        case CellType.FORMULA:
                                            switch (row.GetCell(j).CachedFormulaResultType)
                                            {
                                                case CellType.STRING:
                                                    string strFORMULA = row.GetCell(j).StringCellValue;
                                                    if (strFORMULA != null && strFORMULA.Length > 0)
                                                    {
                                                        dataRow[j] = strFORMULA.ToString();
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = null;
                                                    }
                                                    break;
                                                case CellType.NUMERIC:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                    break;
                                                case CellType.BOOLEAN:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                    break;
                                                case CellType.ERROR:
                                                    dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                    break;
                                                default:
                                                    dataRow[j] = "";
                                                    break;
                                            }
                                            break;
                                        default:
                                            dataRow[j] = "";
                                            break;
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                Console.WriteLine(exception.ToString());
                            }
                        }
                        table.Rows.Add(dataRow);
                    }
                    catch (Exception exception)
                    {

                        Console.WriteLine(exception.ToString());
                    }
                }
            }
            catch (Exception exception)
            {
                //wl.WriteLogs(exception.ToString());
            }
            return table;
        }
 

    }
}
