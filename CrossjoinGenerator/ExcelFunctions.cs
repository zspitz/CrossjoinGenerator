using ADODB;
using NetOffice.ExcelApi;
using Util;

namespace CrossjoinGenerator;
public static class ExcelFunctions {
    private static Application getApp() => new Application {
        Visible = true,
        UserControl = true
    };
    public static string WriteFinal(Recordset rst, string filename) {
        using var excelApp = getApp();

        Workbook book;
        try {
            book = excelApp.Workbooks.Open(filename);
        } catch (Exception) {
            return "Book not found";
        }

        // Wrapping this in a separate method to ensure the COM references are released, enabling deletion
        forceDeleteWorkheet(book);

        var sheet = (Worksheet)book.Worksheets.Add();
        sheet.Name = "Final";
        sheet.DisplayRightToLeft = true;
        sheet.Range("A1").CopyFromRecordset(rst);
        rst.Close();
        rst = null!;

        sheet.Rows[1].Insert();
        sheet.Range("A1:I1").Value = new object[,] {
            {
                "שם", "כיתה נוכחית", "כיתה חדשה", "סוג פריט", "סידורי", "פריט", "מחיר", "כמות", @"סה""כ"
            }
        };
        var lastRow = sheet.UsedRange.Rows.Count;
        sheet.Range($"I2:I{lastRow}").Formula = "=G2*H2";

        sheet.Columns["A:I"].AutoFit();
        return "";
    }

    private static void forceDeleteWorkheet(Workbook book) {
        try {
            using var sheet = (Worksheet)book.Worksheets["Final"];
            using var app = book.Application;
            var displayAlerts = app.DisplayAlerts;
            app.DisplayAlerts = false;
            sheet.Delete();
            app.DisplayAlerts = displayAlerts;
            app.DisposeChildInstances();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        } catch { }
    }

    public static void OpenBook(string filename) {
        using var excelApp = getApp();
        try {
            excelApp.Workbooks.Open(filename, false, false);
        } catch { }
    }

    public static readonly Dictionary<string, List<string>> BookStructure = new() {
        {"Students", ["Name1","Name2","CurrentGrade"] },
        {"Grades", ["CurrentGrade","NewGrade"] },
        {"Items", ["NewGrade","Item","Price","Type","Order"] }
    };

    public static void GenerateTemplate() {
        using var excelApp = getApp();
        var book = excelApp.Workbooks.Add();
        BookStructure.ForEach((kvp, index) => {
            var (sheetname, fields) = kvp;
            var sheet = (Worksheet)book.Worksheets[index + 1];
            sheet.Name = sheetname;
            sheet.DisplayRightToLeft = true;
            fields.ForEach((field, index) => {
                sheet.Cells[1, index + 1].Value2 = field;
            });
        });
    }
}
