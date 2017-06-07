using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Printer.Info;
using LinqToExcel;
using System.Threading;

namespace ExcelImport
{
    public class AsyncInfoUploader
    {
        private InfoManipulator manipulator = new InfoManipulator();
        private readonly string[] mappers;
        private PrinterCategory GetCagegoryFromSheetName(string sheet)
        {
            switch (sheet)
            {
                case "I CLASS":
                    return PrinterCategory.IClass;
                case "M CLASS":
                    return PrinterCategory.MClass;
                default:
                    throw new ArgumentNullException(sheet);
            }
        }
        public string Path { get; set; }
        public AsyncInfoUploader()
        {
            mappers = new string[] {"I CLASS","M CLASS" };
        }
        public async Task<int> UploadExistSheetsAsync(CancellationToken ct,IProgress<double> progress=null)
        {
            var excelFile = new ExcelQueryFactory(Path);
            var sheets = from s in excelFile.GetWorksheetNames()
                         where mappers.Contains(s)
                         select s;
            int count = sheets.Count();
            int j = 0;
            return await Task.Run<int>(()=>{
                foreach (string sheet in sheets)
                {
                    j++;
                    PrinterCategory category = GetCagegoryFromSheetName(sheet);
                    manipulator.DeleteAllInfo(category);
                    var query = excelFile.Worksheet(sheet);
                    int entry_count = query.Count();
                    int i = 0;
                    query.ToList().ForEach((a) => {
                        PrinterInfo info = new PrinterInfo();
                        info.CID = a["CID"];
                        info.AgencyLabel = a["Agency"];
                        info.PackageLabel = a["Package"];
                        manipulator.InsertPrinterInfo(info, category);
                        ct.ThrowIfCancellationRequested();
                        if (progress != null) { progress.Report(((double)++i / entry_count)*j/count); }
                    });
                }
                return j;
            });
        }                
    }
}