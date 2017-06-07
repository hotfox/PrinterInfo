using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Printer.Info;
using System.Threading;
using LinqToExcel;
using System.ComponentModel;

namespace ExcelImport
{
    public class BakgroundInfoUploader:BackgroundWorker
    {
        public string Path { get; set; }
        public PrinterCategory PrinterCategory { get; set; }
        public BakgroundInfoUploader()
        {
            manipulator = new InfoManipulator();
            DoWork += ParseAndUpload;
            WorkerReportsProgress = true;
        }
        public void ParseAndUpload()
        {
            manipulator.DeleteAllInfo(PrinterCategory);
            var excelFile = new ExcelQueryFactory(Path);
            var sheets = excelFile.GetWorksheetNames();
            if (sheets.Contains("M Class CID LIST1"))
            {

            }
            if (sheets.Contains("I ClASS"))
            {

            }
        }
        private InfoManipulator manipulator;
        private void ParseAndUpload(object sender, DoWorkEventArgs e)
        {
            ParseAndUpload();
        }
        private void UploadWorkSheet(ExcelQueryFactory excelFile,string name,PrinterCategory category,int progress)
        {
            var getData = from a in excelFile.Worksheet(name)
                          select a;
            int count = 0;
            int max_count = getData.Count();
            foreach (var a in getData)
            {
                PrinterInfo info = new PrinterInfo();
                info.CID = a["CID"];
                info.AgencyLabel = a["Agency"];
                info.PackageLabel = a["Package"];
                manipulator.InsertPrinterInfo(info, category);
                ++count;
                double p = (double)count / max_count;
                int percent = (int)(p * 100.0);
                ReportProgress(percent);
            }
        }
    }
}
