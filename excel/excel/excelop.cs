using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Printing;//using System.Drawing.Printing ;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace excel
{
    public class excelop
    {
        public Application app;
        public Workbooks wbks;
        public Workbook wbk;
        public Sheets shs;
        public Worksheet wsh;
        public object missig = Missing.Value;
        public Range range;
        public string excelName;
        public excelop(string _excelName)
        {
            init(_excelName);
            //  killExcel();
            openExcel(_excelName);
            setFormat();
            printExcel();
            closeExcel();
            // killExcel();
            // DeleteSheet(@"C:\Users\cl\Desktop\tx.xlsx");
            // createExcel(_excelName);
           // printExcel();
           // GetPrintTicketFromPrinter();
        }
        #region
        /// <summary>
        /// 删除指定的工作表
        /// </summary>
        /// <param name="ExcelName"></param>
        /// 
        public void init(string _excelName)
        {
            app = new Application();
            excelName = _excelName;
        }
        public void deleteSheet(string _excelName)
        {
            //创建 Excel对象
            Application App = new Application();
            //获取缺少的object类型值
            object missing = Missing.Value;
            //打开指定的Excel文件
            Workbook openwb = App.Workbooks.Open(_excelName, missing, missing, missing, missing,
                missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            //删除指定的工作表
            //  Console.WriteLine("请输入你要删除的工作表：");
            string sheetName = "xiaomi";
            ((Worksheet)openwb.Worksheets[sheetName]).Delete();
            Console.WriteLine("删除成功！");
            App.DisplayAlerts = false;//不现实提示对话框
            // Microsoft.Office.Interop.Excel.Range xlsColumns = (Microsoft.Office.Interop.Excel.Range)worksheet.Columns[2, miss];
            // celLrangE = worksheet.Columns;
            // columnCount = celLrangE.Count;
            //   xlsColumns.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, columnCount);
            openwb.Save();//保存工作表
            App.Visible = true;//显示Excel
            openwb.Close(false, missing, missing);//关闭工作表
            //创建进程对象
            Process[] ExcelProcess = Process.GetProcessesByName("Excel");
            //关闭进程
            foreach (Process p in ExcelProcess)
            {
                p.Kill();
            }

        }
        public void createExcel(string _excelName)
        {
            Application excel;
            Workbook worKbooK;
            Worksheet worksheet;
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            worKbooK = excel.Workbooks.Add(Type.Missing);
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
            // worksheet.Name = "StudentRepoertCard";
            // worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 8]].Merge();
            //worksheet.Cells[1, 1] = "Student Report Card";
            worksheet.Cells.Font.Size = 15;
            worKbooK.SaveAs(_excelName);
            worKbooK.Close();
            excel.Quit();
        }
        public void openExcel(string _excelName)
        {

            wbk = app.Workbooks.Open(_excelName);
            wbks = app.Workbooks;
            wsh = (Microsoft.Office.Interop.Excel.Worksheet)wbk.ActiveSheet;
            app.DisplayAlerts = false;

            Console.WriteLine(wsh.Name);

        }

        public void killExcel()
        {
            //创建进程对象
            Process[] ExcelProcess = Process.GetProcessesByName("Excel");
            //关闭进程
            foreach (Process p in ExcelProcess)
            {
                p.Kill();
            }
        }
        public void closeExcel()
        {
            wbk.Save();
            wbk.Close();
            wbks.Close();
            app.Quit();
        }
        public void setFormat()
        {
            char Colum1 = 'A';
            string Colum;
            string ColumT;
            int i;
            for (i = 0; i <= 25; i++)
            {
                Colum1 = (char)(Colum1 + i);
                ColumT = Colum1.ToString();
                Colum = ColumT + ':' + ColumT;
                range = wsh.get_Range(Colum, System.Type.Missing);
                range.EntireColumn.ColumnWidth = 0.7;
                Colum1 = 'A';
            }
            //  range = wsh.get_Range(Colum, System.Type.Missing);
            // range.EntireColumn.ColumnWidth = 0.65;

            for (i = 0; i <= 25; i++)
            {
                Colum1 = (char)(Colum1 + i);
                ColumT = "A" + Colum1.ToString();
                Colum = ColumT + ':' + ColumT;
                range = wsh.get_Range(Colum, System.Type.Missing);
                range.EntireColumn.ColumnWidth = 0.7;
                Colum1 = 'A';
            }
            for (i = 0; i <= 18; i++)
            {
                Colum1 = (char)(Colum1 + i);
                ColumT = "B" + Colum1.ToString();
                Colum = ColumT + ':' + ColumT;
                range = wsh.get_Range(Colum, System.Type.Missing);
                range.EntireColumn.ColumnWidth = 0.7;
                Colum1 = 'A';
            }
            for (i = 1; i <= 30; i++)
            {
                range = (Range)wsh.Rows[i, Missing.Value];
                range.RowHeight = 21;
            }
            range = wsh.get_Range("N:N", System.Type.Missing);
            range.EntireColumn.ColumnWidth = 4.25;

            range = (Range)wsh.Rows[4, Missing.Value];
            range.RowHeight = 27.5;
            merge();



        }
        public void merge()
        {
            int row, colum, i;
            row = 1;
            colum = 1;
            for (i = 0; i < 3; i++)
            {
                wsh.Range[wsh.Cells[row, colum], wsh.Cells[row, colum + 10]].Merge();
                row++;
            }
            row = 1;
            colum = 12;
            wsh.Range[wsh.Cells[row, colum], wsh.Cells[row, colum + 2]].Merge();
            colum = 15;
            wsh.Range[wsh.Cells[row, colum], wsh.Cells[row, colum + 1]].Merge();
            colum = 17;
            wsh.Range[wsh.Cells[row, colum], wsh.Cells[row, colum + 6]].Merge();
            colum = 24;
            wsh.Range[wsh.Cells[row, colum], wsh.Cells[row, colum + 17]].Merge();
            colum = 42;
            wsh.Range[wsh.Cells[row, colum], wsh.Cells[row, colum + 10]].Merge();
            colum = 53;
            wsh.Range[wsh.Cells[row, colum], wsh.Cells[row, colum + 15]].Merge();
            ///////////////////////
            row = 2;
            for (i = 0; i < 2; i++)
            {
                wsh.Range[wsh.Cells[row, 12], wsh.Cells[row, 41]].Merge();
                wsh.Range[wsh.Cells[row, 42], wsh.Cells[row, 52]].Merge();
                wsh.Range[wsh.Cells[row, 53], wsh.Cells[row, 68]].Merge();
                row++;
            }
            row = 4;
            for (i = 0; i < 10; i++)
            {
                wsh.Range[wsh.Cells[row, 1], wsh.Cells[row, 14]].Merge();
                wsh.Range[wsh.Cells[row, 15], wsh.Cells[row, 32]].Merge();
                wsh.Range[wsh.Cells[row, 33], wsh.Cells[row, 50]].Merge();
                wsh.Range[wsh.Cells[row, 51], wsh.Cells[row, 68]].Merge();
                row++;
            }
            row = 14;
            for (i = 0; i < 6; i++)
            {
                wsh.Range[wsh.Cells[row, 1], wsh.Cells[row, 4]].Merge();
                wsh.Range[wsh.Cells[row, 5], wsh.Cells[row, 14]].Merge();
                wsh.Range[wsh.Cells[row, 15], wsh.Cells[row, 23]].Merge();
                wsh.Range[wsh.Cells[row, 24], wsh.Cells[row, 32]].Merge();
                wsh.Range[wsh.Cells[row, 33], wsh.Cells[row, 41]].Merge();
                wsh.Range[wsh.Cells[row, 42], wsh.Cells[row, 50]].Merge();
                wsh.Range[wsh.Cells[row, 51], wsh.Cells[row, 59]].Merge();
                wsh.Range[wsh.Cells[row, 60], wsh.Cells[row, 68]].Merge();
                row++;
            }
            wsh.Range[wsh.Cells[14, 1], wsh.Cells[19, 1]].Merge();
            row = 19;
            wsh.Range[wsh.Cells[row, 15], wsh.Cells[row, 32]].Merge();
            wsh.Range[wsh.Cells[row, 33], wsh.Cells[row, 50]].Merge();
            wsh.Range[wsh.Cells[row, 51], wsh.Cells[row, 68]].Merge();
            row = 20;
            for (i = 0; i < 4; i++)
            {
                wsh.Range[wsh.Cells[row, 1], wsh.Cells[row, 14]].Merge();
                wsh.Range[wsh.Cells[row, 15], wsh.Cells[row, 32]].Merge();
                wsh.Range[wsh.Cells[row, 33], wsh.Cells[row, 50]].Merge();
                wsh.Range[wsh.Cells[row, 51], wsh.Cells[row, 68]].Merge();
                row++;
            }
            colum = 1;
            wsh.Range[wsh.Cells[row, colum], wsh.Cells[row, colum + 10]].Merge();
            wsh.Range[wsh.Cells[row, 12], wsh.Cells[row, 68]].Merge();
            printExcel();

        }
        public void printExcel()
        {
            var printServer = new PrintServer();

            //获取全部打印机
            var printQueueCol = printServer.GetPrintQueues();
            foreach (PrintQueue printer in printQueueCol)
            {
                Console.WriteLine("tThe shared printer " + printer.Name + " is located at " + printer.Location + "\n");
            }
            //获取默认打印机
            var defaultPrintQue = LocalPrintServer.GetDefaultPrintQueue();
            var ticket = defaultPrintQue.DefaultPrintTicket;
            ticket.PageOrientation = PageOrientation.Landscape;
            string str = ticket.InputBin.ToString();
           // ticket.InputBin = InputBin.AutoSheetFeeder;
            var printCapability = defaultPrintQue.GetPrintCapabilities();
            ReadOnlyCollection<InputBin> binBox = printCapability.InputBinCapability;
            string name = defaultPrintQue.Name.ToString();
            //   AutoSheetFeeder
            //foreach(InputBin box in binBox)
            //{
            //    Console.WriteLine(box.ToString());
            //}
            //LocalPrintServer ps = new LocalPrintServer();
            //// Get the default print queue
            //PrintQueue pq = ps.DefaultPrintQueue;
            //Console.WriteLine(pq.ToString());
            //ps.Commit();
            // foreach (PrintQueue printer in pq)
            //{
            //    Console.WriteLine("tThe shared printer " + printer.Name + " is located at " + printer.Location +"\n");
            //}
            // Get an XpsDocumentWriter for the default print queue
            //XpsDocumentWriter xpsdw = PrintQueue.CreateXpsDocumentWriter(pq);
            //return xpsdw;
            //------------------------打印页面相关设置--------------------------------
            // wsh.PageSetup.PaperSize = XlPaperSize.xlPaperA4;//纸张大小
            // wsh.PageSetup.Orientation = XlPageOrientation.xlPortrait;//页面横向
            //wsh.PageSetup.Zoom = 75; //打印时页面设置,缩放比例百分之几
            // wsh.PageSetup.Zoom = false; //打印时页面设置,必须设置为false,页高,页宽才有效
            // wsh.PageSetup.FitToPagesWide = 1; //设置页面缩放的页宽为1页宽
            //// wsh.PageSetup.FitToPagesTall = false; //设置页面缩放的页高自动
            //wsh.PageSetup.LeftHeader = "Nigel";//页面左上边的标志
            //wsh.PageSetup.CenterFooter = "第 &P 页，共 &N 页";//页面下标
            //wsh.PageSetup.PrintGridlines = true; //打印单元格网线
            // wsh.PageSetup.TopMargin = 1.5 / 0.035; //上边距为2cm（转换为in）
            //  wsh.PageSetup.BottomMargin = 1.5 / 0.035; //下边距为1.5cm
            // wsh.PageSetup.LeftMargin = 2 / 0.035; //左边距为2cm
            // wsh.PageSetup.RightMargin = 2 / 0.035; //右边距为2cm
            //  wsh.PageSetup.CenterHorizontally = true; //文字水平居中
            //------------------------打印页面设置结束--------------------------------
            // app.Visible = true;
            // wbk.PrintPreview(); //打印预览
            //wbk.PrintOutEx();
            wbk.PrintOutEx(); //直接打印
            //worksBook.Close(); //关闭工作空间
            //ExcelApp.Quit(); //退出程序
            //wsh.SaveAs(filePath); //另存表        }
        #endregion
        }
        // ---------------------- GetPrintTicketFromPrinter -----------------------
        /// <summary>
        ///   Returns a PrintTicket based on the current default printer.</summary>
        /// <returns>
        ///   A PrintTicket for the current local default printer.</returns>
        private PrintTicket GetPrintTicketFromPrinter()
        {
            PrintQueue printQueue = null;

            LocalPrintServer localPrintServer = new LocalPrintServer();

            // Retrieving collection of local printer on user machine
            PrintQueueCollection localPrinterCollection =
                localPrintServer.GetPrintQueues();

            System.Collections.IEnumerator localPrinterEnumerator =
                localPrinterCollection.GetEnumerator();

            if (localPrinterEnumerator.MoveNext())
            {
                // Get PrintQueue from first available printer
                printQueue = (PrintQueue)localPrinterEnumerator.Current;
            }
            else
            {
                // No printer exist, return null PrintTicket
                return null;
            }

            // Get default PrintTicket from printer
            PrintTicket printTicket = printQueue.DefaultPrintTicket;

            PrintCapabilities printCapabilites = printQueue.GetPrintCapabilities();

            // Modify PrintTicket
            if (printCapabilites.CollationCapability.Contains(Collation.Collated))
            {
                printTicket.Collation = Collation.Collated;
            }

            if (printCapabilites.DuplexingCapability.Contains(
                    Duplexing.TwoSidedLongEdge))
            {
                printTicket.Duplexing = Duplexing.TwoSidedLongEdge;
            }

            if (printCapabilites.StaplingCapability.Contains(Stapling.StapleDualLeft))
            {
                printTicket.Stapling = Stapling.StapleDualLeft;
            }

            return printTicket;
        }// end:GetPrintTicketFromPrinter()    }
    }
}
