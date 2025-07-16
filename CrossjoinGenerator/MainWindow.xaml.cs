using Ookii.Dialogs.Wpf;
using System.Windows;
using Util;
using static CrossjoinGenerator.ExcelFunctions;
using static CrossjoinGenerator.RecordsetFunctions;

namespace CrossjoinGenerator;

public partial class MainWindow : Window {
    public MainWindow() {
        InitializeComponent();

        DataContext = new MainViewModel();

        //chooseFile.Click += (s, e) => {
        //    var dlg = new VistaOpenFileDialog {
        //        Filter = "קבצי Excel|*.xlsx"
        //    };
        //    if ((bool)dlg.ShowDialog(this)!) {
        //        filename.Text = dlg.FileName;
        //    }
        //};

        //start.Click += (s, e) => {
        //    message.Text = "";
        //    invalidData.ItemsSource = null;

        //    Filename = filename.Text;

        //    var msg1 = testSQL();
        //    if (!msg1.IsNullOrWhitespace()) {
        //        message.Text = msg1;
        //        return;
        //    }

        //    var (msg, sqlFrom) = validate();
        //    if (!msg.IsNullOrWhitespace()) {
        //        message.Text = msg;
        //        var rst = GetRst($"SELECT * {sqlFrom}");
        //        var dt = rst.ToDataTable();
        //        ReleaseRst(ref rst);
        //        invalidData.ItemsSource = dt.DefaultView;
        //        return;
        //    }

        //    writeFinal();
        //};

        //openFile.Click += (s,e) => OpenBook(filename.Text);
    }

    private readonly string sqlFrom =
@"FROM ([Students$] AS Students
INNER JOIN [Grades$] AS Grades ON Students.CurrentGrade = Grades.CurrentGrade)
INNER JOIN [Items$] AS Items ON Grades.NewGrade = Items.NewGrade";

    private readonly string[] sheets = [
        "Students",
        "Grades",
        "Items"
    ];

    private readonly (string, string)[] fields = [
        ("Students","Name1"),
        ("Students","Name2"),
        ("Students","CurrentGrade"),
        ("Grades","CurrentGrade"),
        ("Grades","NewGrade"),
        ("Items","NewGrade"),
        ("Items","Item"),
        ("Items","Price"),
        ("Items","Type"),
        ("Items","Order")
    ];

    private string testSQL() {
        foreach (var sheet in sheets) {
            try {
                TestSingleSql($"SELECT * FROM [{sheet}$]");
            } catch (Exception) {
                return $"חסר דף {sheet} בקובץ Excel";
            }
        }

        foreach (var (sheet, field) in fields) {
            try {
                TestSingleSql($"SELECT [{field}] FROM [{sheet}$]");
            } catch (Exception) {
                return $"חסר עמודה {field} בדף {sheet}";
            }
        }

        try {
            TestSingleSql($"SELECT * {sqlFrom}");
        } catch (Exception ex) {
            return "שגיאה בצירוף עמודות" + ex.Message;
        }

        foreach (var (sheet, field) in fields) {
            try {
                TestSingleSql($"SELECT {sheet}.{field} {sqlFrom}");
            } catch (Exception) {
                return $"חסר עמודה {field} בדף {sheet}";
            }
        }

        return "";
    }

    private readonly List<(string, string)> validations = [
        
        // השם המלא יכול להיות או ב-Name1 או ב-Name2, או בשניהם
        // הבעיה היא אם אין שם ויש כיתה נוכחית
        (
            "חסר שם התלמיד",
            "FROM [Students$] WHERE Name1 IS NULL AND Name2 IS NULL AND CurrentGrade IS NOT NULL"
        ),

        // warning
        (
            "תלמידים עם כיתה נוכחית לא תקינה",
            @"
                FROM [Students$] AS Students
                LEFT JOIN [Grades$] AS Grades ON Students.CurrentGrade = Grades.CurrentGrade
                WHERE Grades.CurrentGrade IS NULL
            "
        ),

        // warning
        (
            "כיתה ללא תלמידים",
            @"
                FROM [Grades$] AS Grades
                LEFT JOIN [Students$] AS Students ON Grades.CurrentGrade = Students.CurrentGrade
                WHERE Students.CurrentGrade IS NULL
            "
        ),

        // warning
        (
            "כתה ללא פריטים",
            @"
                FROM [Grades$] AS Grades
                LEFT JOIN [Items$] AS Items ON Grades.NewGrade = Items.NewGrade
                WHERE Items.NewGrade IS NULL
            "
        ),

        (
            "פריטים עם כיתה לא תקין",
            @"
                FROM [Items$] AS Items
                LEFT JOIN [Grades$] AS Grades ON Items.NewGrade = Grades.NewGrade
                WHERE Item IS NOT NULL AND Grades.NewGrade IS NULL
            "
        )
    ];

    private (string, string) validate() {
        foreach (var (message, sqlFrom) in validations) {
            var success = validateSingle(sqlFrom);
            if (!success) {
                return (message, sqlFrom);
            }
        }
        return ("", "");
    }

    // Returns true if the validation was successful
    private bool validateSingle(string sqlFrom) {
        var rst = GetRst($"SELECT COUNT(*) {sqlFrom}");
        int count = rst.Fields[0].Value;
        ReleaseRst(ref rst);
        return count == 0;
    }

    private void writeFinal() {
        var sqlFinal = $@"
SELECT Students.Name1 & "" "" & Students.Name2, Students.CurrentGrade, Grades.NewGrade, Items.Type, Items.Order, Items.Item, Items.Price, NULL AS Qty
{sqlFrom}
ORDER BY Students.CurrentGrade, Grades.NewGrade, Students.Name1, Students.Name2, Items.Order";

        var rst = GetRst(sqlFinal);
        WriteFinal(rst, filename.Text);
        ReleaseRst(ref rst);
    }
}
