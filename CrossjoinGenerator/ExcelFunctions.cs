using ClosedXML.Excel;
using NetOffice.ExcelApi;
using Util;

namespace CrossjoinGenerator;
public static class ExcelFunctions {
    private static Application getApp() => new() {
        Visible = true,
        UserControl = true
    };

    public static void WriteFinal2(System.Data.DataTable dt, string filename) {
        using var workbook = new XLWorkbook(filename);
        if (workbook.TryGetWorksheet("Final", out var finalSheet)) {
            finalSheet.Delete();
        }

        var sheet = workbook.AddWorksheet("Final", 1); // TODO is this position one-based or zero-based?
        sheet.RightToLeft = true;

        sheet.Cell(1, 1).InsertTable(dt, false);

        var headers = new[] {
            "שם", "כיתה נוכחית", "כיתה חדשה", "סוג פריט", "סידורי", "פריט", "מחיר", "כמות", @"סה""כ"
        };
        for (var i = 0; i < headers.Length; i++) {
            sheet.Cell(1, i + 1).Value = headers[i]; // TODO are cell references one-based or zero-based?
        }

        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        for (var row = 2; row <= lastRow; row++) {
            sheet.Cell(row, 8).Value = Blank.Value;
            sheet.Cell(row, 9).FormulaA1 = $"=G{row}*H{row}";
        }

        sheet.Columns(1, 9).AdjustToContents();

        workbook.Save(new() {
            EvaluateFormulasBeforeSaving = true,
            ValidatePackage = true
        });

        OpenBook(filename);
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
