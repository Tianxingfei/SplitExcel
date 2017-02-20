using System;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Collections.Generic;
using HelperExtend;
namespace SplitExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string modelExlPath = System.Configuration.ConfigurationManager.AppSettings["modelExlPath"];
            //需要添加 Microsoft.Office.Interop.Excel引用 

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            if (app == null)
            {
                Console.WriteLine("服务器上缺少Excel组件，需要安装Office软件。");
                return;
            }
            object missing = Missing.Value;
            Workbook wb = null;
            Workbook Sourceworkbook = null;
            Worksheet Sourceworksheet = null;
            List<string> filename = new List<string>();
            int column1 = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["column1"]);
            int column2 = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["column2"]);

            int column1sheet = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["column1sheet"]);
            int column2sheet = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["column2sheet"]);

            int Maxcolumn = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Maxcolumn"]);
            int startrow = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["startrow"]);
            string Writecolumn = System.Configuration.ConfigurationManager.AppSettings["Writecolumn"];
            string Classification = System.Configuration.ConfigurationManager.AppSettings["Classification"];
            string ClassificationValue = System.Configuration.ConfigurationManager.AppSettings["ClassificationValue"];

            string sheetNum = System.Configuration.ConfigurationManager.AppSettings["sheetNum"];

            String SourceExlPath = System.Configuration.ConfigurationManager.AppSettings["SourceExlPath"];
            string[] writecolumn = Writecolumn.Split(',');
            List<int> writecolumnlist = new List<int>();

            string PassWord = System.Configuration.ConfigurationManager.AppSettings["PassWord"];

            //加载模板,打开excel

            //try
            //{

            Console.WriteLine("读取文件excel");
            //READ template
            Console.WriteLine(modelExlPath );
            wb = app.Workbooks.Open(modelExlPath, false, missing, missing, PassWord, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            app.Visible = false;
            app.UserControl = false;
            app.DisplayAlerts = false;
            app.AlertBeforeOverwriting = false;

            //1、获取数据。
            //获取SourceExcle的数据
            //Sourceworkbook = app.Workbooks.Open(SourceExlPath, false, missing, missing, PassWord, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            //Sourceworksheet = (Worksheet)Sourceworkbook.Worksheets.get_Item(2);
            HelperExtend.AsFunctionExcelIO help = new HelperExtend.AsFunctionExcelIO();
            System.Data.DataTable dt = help.ImportExcelToDataTable(SourceExlPath, Convert.ToInt32(sheetNum));
            Console.WriteLine("读取文件excel完毕");

            //获取数据列
            for (int i = 0; i < writecolumn.Count(); i++)
            {
                writecolumnlist.Add(Convert.ToInt32(writecolumn[i]));
            }

            //input excel records
            Console.WriteLine("开始分类保存模板文件");

            //获取保存文件的文件名
            System.Data.DataView dv = dt.DefaultView;

            System.Data.DataTable dtdistinct1 = dv.ToTable(true, "A" + column1.ToString());
            System.Data.DataTable dtdistinct2 = dv.ToTable(true, "A" + column2.ToString());

            for (int i = 0; i < dtdistinct1.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(dtdistinct1.Rows[i][0].ToString()))
                {
                    filename.Add(dtdistinct1.Rows[i][0].ToString());
                }
            }
            for (int i = 0; i < dtdistinct2.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(dtdistinct2.Rows[i][0].ToString()))
                {
                    filename.Add(dtdistinct2.Rows[i][0].ToString());
                }
            }
            Console.WriteLine("导出模板文件数" + filename.Count().ToString());
            //保存模板至每一个Excel中
            for (int j = 0; j < filename.Count(); j++)
            {

                string downExlPath = System.Configuration.ConfigurationManager.AppSettings["downExlPath"] + filename[j] + ".xlsx";
                Console.WriteLine("开始保存模板文件" + (j + 1).ToString() + " /" + filename.Count().ToString() + "   " + filename[j]);
                System.IO.FileInfo fileinfo = new System.IO.FileInfo(downExlPath);
                if (!fileinfo.Exists)
                {
                    wb.SaveAs(downExlPath, Missing.Value, PassWord, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    Console.WriteLine("保存模板" + filename[j]);
                }
            }

            //按顺序释放资源。
            wb.Close(false, modelExlPath, false);
            NAR(wb);


            //打开每一个Excle，并写入数据
            for (int j = 0; j < filename.Count(); j++)
            {
                try
                {
                    //打开Excle
                    string downExlPath = System.Configuration.ConfigurationManager.AppSettings["downExlPath"] + filename[j] + ".xlsx";
                    System.IO.FileInfo fileinfo = new System.IO.FileInfo(downExlPath);
                    if (fileinfo.Length > 454000)
                    {
                        Workbook Workbook = app.Workbooks.Open(downExlPath, false, missing, missing, PassWord, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                        app.Visible = false;

                        //得到WorkSheet对象
                        Worksheet worksheet = (Worksheet)Workbook.Worksheets.get_Item(column1sheet);
                        Worksheet nokiaworksheet = (Worksheet)Workbook.Worksheets.get_Item(column2sheet);
                        Console.WriteLine("打开模板" + filename[j]);
                        //2、写入数据，Excel索引从1开始。
                        int m = 0; int n = 0;
                        System.Data.DataRow[] mycach1 = dt.Select("[A" + column1.ToString() + "]='" + filename[j].ToString() + "'");
                        Console.WriteLine("写入模板" + filename[j]);
                        //for (int p = startrow; p < Sourceworksheet.UsedRange.Rows.Count; p++)
                        if (mycach1.Count() > 0)
                        {
                            for (int z = 0; z < mycach1.Count(); z++)
                            {
                                System.Data.DataRow myrow = mycach1[z];
                                if (myrow[Classification].ToString().ToUpper() == ClassificationValue)
                                {
                                    foreach (int i in writecolumnlist)
                                    {
                                        nokiaworksheet.Cells[startrow + m, i] = myrow[i - 1];// Sourceworksheet.Cells[p, i];
                                    }
                                    m++;
                                }
                                else
                                {
                                    foreach (int i in writecolumnlist)
                                    {
                                        worksheet.Cells[startrow + n, i] = myrow[i - 1];// Sourceworksheet.Cells[p, i];
                                    }
                                    n++;
                                }
                                Console.WriteLine("写入 " + filename[j] + " -- " + (z + 1).ToString() + " row, 共" + mycach1.Count() + "条，已完成" + (j).ToString() + "个文件，共" + filename.Count() + "文件。");
                                #region discard
                                //if (((Range)Sourceworksheet.Cells[p, column1]).Text.ToString() == filename[j] || ((Range)Sourceworksheet.Cells[p, column2]).Text.ToString() == filename[j])
                                //{
                                //    if (((Range)Sourceworksheet.Cells[p, 2]).Text.ToString().ToLower() == "NOKIA".ToLower())
                                //    {
                                //        foreach (int i in writecolumnlist)
                                //        {
                                //            nokiaworksheet.Cells[startrow + m, i] = Sourceworksheet.Cells[p, i];
                                //        }
                                //        m++;
                                //    }
                                //    else
                                //    {
                                //        foreach (int i in writecolumnlist)
                                //        {
                                //            worksheet.Cells[startrow + n, i] = Sourceworksheet.Cells[p, i];
                                //        }
                                //        n++;
                                //    }
                                //}
                                #endregion
                            }
                        }
                        System.Data.DataRow[] mycach2 = dt.Select("[A" + column2.ToString() + "]='" + filename[j].ToString() + "'");
                        //for (int p = startrow; p < Sourceworksheet.UsedRange.Rows.Count; p++)
                        if (mycach2.Count() > 0)
                        {
                            for (int z = 0; z < mycach2.Count(); z++)
                            {
                                System.Data.DataRow myrow = mycach2[z];
                                if (myrow[Classification].ToString().ToUpper() == ClassificationValue)
                                {
                                    foreach (int i in writecolumnlist)
                                    {
                                        nokiaworksheet.Cells[startrow + m, i] = myrow[i - 1];// Sourceworksheet.Cells[p, i];
                                    }
                                    m++;
                                }
                                else
                                {
                                    foreach (int i in writecolumnlist)
                                    {
                                        worksheet.Cells[startrow + n, i] = myrow[i - 1];// Sourceworksheet.Cells[p, i];
                                    }
                                    n++;
                                }
                                #region discard
                                //if (((Range)Sourceworksheet.Cells[p, column1]).Text.ToString() == filename[j] || ((Range)Sourceworksheet.Cells[p, column2]).Text.ToString() == filename[j])
                                //{
                                //    if (((Range)Sourceworksheet.Cells[p, 2]).Text.ToString().ToLower() == "NOKIA".ToLower())
                                //    {
                                //        foreach (int i in writecolumnlist)
                                //        {
                                //            nokiaworksheet.Cells[startrow + m, i] = Sourceworksheet.Cells[p, i];
                                //        }
                                //        m++;
                                //    }
                                //    else
                                //    {
                                //        foreach (int i in writecolumnlist)
                                //        {
                                //            worksheet.Cells[startrow + n, i] = Sourceworksheet.Cells[p, i];
                                //        }
                                //        n++;
                                //    }
                                //}
                                #endregion
                                Console.WriteLine("写入 " + filename[j] + " -- " + (z + 1).ToString() + " row，共" + mycach2.Count() + "条，已完成" + (j).ToString() + "个文件，共" + filename.Count() + "文件。");
                            }
                        }

                        Range range = (Range)worksheet.get_Range(worksheet.Cells[startrow + n, 1], worksheet.Cells[Maxcolumn, column2]);
                        range.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                        Range range1 = (Range)nokiaworksheet.get_Range(nokiaworksheet.Cells[startrow + m, 1], nokiaworksheet.Cells[Maxcolumn, column2]);
                        range1.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                        Console.WriteLine("删除文件多余的行");
                        //3、保存生成的Excel文件。
                        Workbook.SaveAs(downExlPath, Missing.Value, PassWord, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        Console.WriteLine("保存文件" + filename[j] + "成功。");
                        Workbook.Close(missing, missing, missing);
                        NAR(worksheet);
                        NAR(nokiaworksheet);
                        NAR(Workbook);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    Console.WriteLine("程序运行中请不要打开excle，如打开，请关闭excle");
                    continue;
                }
            }

            //按顺序释放资源。
            NAR(Sourceworksheet);
            NAR(Sourceworkbook);
            app.Quit();
            NAR(app);
            Console.WriteLine(System.DateTime.Now.ToShortTimeString() + "脚本运行成功!");
            Console.Write("按任意键退出");
            Console.ReadKey();
            
        }

        private static void NAR(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch
            {
            }
            finally
            {
                o = null;
            }
        }

    }
}
