using ADODB;
using NetOffice.ExcelApi;

namespace CrossjoinGenerator; 
public static class ExcelFunctions {
    public static string WriteFinal(Recordset rst, string filename) {
        using var excelApp = new Application {
            Visible = true,
            UserControl = true
        };

        Workbook book;
        try {
            book = excelApp.Workbooks.Open(filename);
        } catch (Exception) {
            return "Book not found";
        }

        Worksheet? sheet=null;
        try {
            sheet = (Worksheet)book.Worksheets["Final"];
        } catch {}
        if (sheet is not null) {
            var displayAlerts = excelApp.DisplayAlerts;
            excelApp.DisplayAlerts = false;
            sheet.Delete();
            excelApp.DisplayAlerts = displayAlerts;
        }

        sheet=(Worksheet)book.Worksheets.Add();
        sheet.Name = "Final";
        sheet.DisplayRightToLeft = true;
        sheet.Range("A1").CopyFromRecordset(rst);
        rst.Close();
        rst = null;

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

    public static void OpenBook(string filename) {
        using var excelApp = new Application {
            Visible = true,
            UserControl = true
        };
        try {
            excelApp.Workbooks.Open(filename,false, false);
        } catch {}
    }
}
