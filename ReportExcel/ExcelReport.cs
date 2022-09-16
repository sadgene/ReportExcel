using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.Linq;
using System.Data.Linq.Mapping;
using ReportExcel;
using System.Reflection;
using Microsoft.CSharp;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

public class Progress
{
    public const string cnt_Export_Directory = "C:\\Users\\user\\Desktop\\_Export_IPS";

    static void Main()
    {
        StringBuilder strBuilder = new StringBuilder();

        if (!Directory.Exists(cnt_Export_Directory))
        {
            Directory.CreateDirectory(cnt_Export_Directory);
            strBuilder.AppendLine(DateTime.Now.ToString() + " Creating directory " + cnt_Export_Directory);//+Это
        }

        string datasource = @"Server";
        string database = "db";
        string username = "user";
        string password = "password";
        string connString = @"Data Source=" + datasource + ";Initial Catalog=" + database + ";Persist Security Info= no ;User ID = " + username + "; Password = " + password;

        DataContext db = new DataContext(connString);

        try
        {

            GenerateFile(db,"Институт1");
            GenerateFile(db, "Институт2");
            GenerateFile(db, "Институт3");
            GenerateFile(db, "Институт4");
            GenerateFile(db, "Институт5");

            db.Dispose();
        }
        catch(Exception ex)
        {
            strBuilder.AppendLine(ex.Message);
            strBuilder.AppendLine(DateTime.Now.ToString() + " Export task ended");
            File.AppendAllText(cnt_Export_Directory + "\\export_projects.log", strBuilder.ToString(), Encoding.UTF8);
        }
      

   }
    static void GenerateFile(DataContext db, string parms)
    {
        int i = 1;
        int j = 1;

        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;
        xlApp = new Excel.Application();
        xlWorkBook = xlApp.Workbooks.Add(misValue);
        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

        dynamic header = xlWorkSheet.Range["A01", "I1"];
        header.Interior.Color = 0x926036;
        header.Font.ColorIndex = 2;
        header.Font.Size = 11;
        header.Font.Bold = true;
        header.Font.Name = "Calibri";
        header.Borders.LineStyle = 1;
        header.Borders.Weight = 2;
        header.Borders.ColorIndex = 1;
        header.ColumnWidth = 16;
        header.WrapText = true;

        header.CellS().Value = new String[]
        {
               "Идентификатор проекта",
               "Институт",
               "Наряд-заказ",
               "Название договора",
               "Заказчик",
               "Плановый прогресс",
               "Фактический прогресс",
               "Отставание",
               "Менеджер РКП"
        };

        var All = db.GetTable<PROGRESS>().Where(k => k.NAME == parms).OrderBy(s => s.Lags);
        foreach (var c in All)
        {

            i++;


            xlWorkSheet.Cells[i, j] = c.PROJ_REL_ID;
            xlWorkSheet.Cells[i, j + 1] = c.NAME;

            if (c.Lags < 0)
                xlWorkSheet.Cells[i, j + 2].Interior.Color = 5287936;
            if (c.Lags > 0 & c.Lags < 10)
                xlWorkSheet.Cells[i, j + 2].Interior.Color = 5296274;
            if (c.Lags > 10 & c.Lags < 30)
                xlWorkSheet.Cells[i, j + 2].Interior.Color = 65535;
            if (c.Lags > 30 & c.Lags < 50)
                xlWorkSheet.Cells[i, j + 2].Interior.Color = 49407;
            if (c.Lags >= 50 & c.Lags <= 100)
                xlWorkSheet.Cells[i, j + 2].Interior.Color = 255;


            xlWorkSheet.Cells[i, j + 2] = c.STRING_PROJREL_ORDER;
            xlWorkSheet.Cells[i, j + 3] = c.PROJ_REL_CAPTION;
            xlWorkSheet.Cells[i, j + 3].ColumnWidth = 165;
            xlWorkSheet.Cells[i, j + 3].WrapText = true;
            xlWorkSheet.Cells[i, j + 4] = c.PROJ_REL_CUSTOMER;
            xlWorkSheet.Cells[i, j + 4].ColumnWidth = 28;
            xlWorkSheet.Cells[i, j + 5] = c.I_PLAN_PROGRESS;
            xlWorkSheet.Cells[i, j + 6] = c.I_FACT_PROGRESS;
            if (c.I_PLAN_PROGRESS > c.I_FACT_PROGRESS)
                xlWorkSheet.Cells[i, j + 7] = c.Lags;
            else
                xlWorkSheet.Cells[i, j + 7] = 0;
            xlWorkSheet.Cells[i, j + 8] = c.USER_NAME;
        }

        xlApp.DisplayAlerts = false;
        xlWorkBook.DoNotPromptForConvert = true;
        xlWorkBook.CheckCompatibility = false;
        xlWorkBook.SaveAs($"C:\\Users\\user\\Desktop\\ExcelReports\\{parms}.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);

        xlWorkBook.Close(true, misValue, misValue);
        xlApp.Quit();
    }
}


namespace ReportExcel
{      

    [Table(Name = "PROGRESS")]
    public class PROGRESS
    {

        [Column]
        public double? PROJ_REL_ID { get; set; }       
        [Column]
        public string NAME { get; set; }
        [Column]
        public string STRING_PROJREL_ORDER { get; set; }     
        [Column]
        public string PROJ_REL_CAPTION { get; set; }
        [Column]
        public string PROJ_REL_CUSTOMER { get; set; }
        [Column]
        public decimal? I_PLAN_PROGRESS { get; set; }
        [Column]
        public decimal? I_FACT_PROGRESS { get; set; }
        [Column]
        public decimal? Lags { get; set; }
        [Column]
        public string USER_NAME { get; set; }
    }
}
