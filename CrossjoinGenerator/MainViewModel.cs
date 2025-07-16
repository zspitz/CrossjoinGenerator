using Ookii.Dialogs.Wpf;
using System.Windows;
using Util;
using static CrossjoinGenerator.ExcelFunctions;
using static CrossjoinGenerator.ProcessState;
using static CrossjoinGenerator.RecordsetFunctions;
using static Util.Functions;

namespace CrossjoinGenerator;

public class MainViewModel : ViewModelBase {
    public RelayCommand GenerateTemplate = new(o => GenerateTemplate());

    public RelayCommand ChooseFile;
    public RelayCommand ProcessFile;
    public RelayCommand EditFile;

    private string filename = "";
    public string Filename {
        get => filename;
        set => NotifyChanged(ref filename, value);
    }

    private bool isRunning = false;
    public bool IsRunning {
        get => isRunning;
        set => NotifyChanged(ref isRunning, value);
    }

    private ProcessState? processState = null;
    public ProcessState? ProcessState {
        get => processState;
        set {
            if (value != null) {
                if (processState == Error) { return; }
                if (processState == Warning && value != Error) { return; }
            }
            NotifyChanged(ref processState, value);
        }
    }

    private string errorMessage = "";
    public string ErrorMessage {
        get => errorMessage;
        set => NotifyChanged(ref errorMessage, value);
    }

    public List<DataCheck> DataChecks { get; } = [
        new(
            "חסר שם התלמיד",
            "SELECT * FROM [Students$] WHERE Name1 IS NULL AND Name2 IS NULL AND CurrentGrade IS NOT NULL",
            false
        ),
        new(
            "תלמידים עם כיתה נוכחית לא תקינה",
            @"
                SELECT *
                FROM [Students$] AS Students
                LEFT JOIN [Grades$] AS Grades ON Students.CurrentGrade = Grades.CurrentGrade
                WHERE Grades.CurrentGrade IS NULL
            ",
            true
        ),
        new(
            "כיתה ללא תלמידים",
            @"
                SELECT *
                FROM [Grades$] AS Grades
                LEFT JOIN [Students$] AS Students ON Grades.CurrentGrade = Students.CurrentGrade
                WHERE Students.CurrentGrade IS NULL
            ",
            true
        ),
        new(
            "כתה ללא פריטים",
            @"
                SELECT *
                FROM [Grades$] AS Grades
                LEFT JOIN [Items$] AS Items ON Grades.NewGrade = Items.NewGrade
                WHERE Items.NewGrade IS NULL
            ",
            true
        ),
        new(
            "פריטים עם כיתה לא תקין",
            @"
                SELECT *
                FROM [Items$] AS Items
                LEFT JOIN [Grades$] AS Grades ON Items.NewGrade = Grades.NewGrade
                WHERE Item IS NOT NULL AND Grades.NewGrade IS NULL
            ",
            false
        )
    ];

    public MainViewModel() {
        bool canExecute(object? o) => !Filename.IsNullOrWhitespace();
        ProcessFile = new(processFile, canExecute);
        EditFile = new(o => OpenBook(Filename), canExecute);

        ChooseFile = new(chooseFile, o => !IsRunning);
    }

    private void chooseFile(object? o) {
        var dlg = new VistaOpenFileDialog {
            Filter = "קבצי Excel|*.xlsx"
        };
        if ((bool)dlg.ShowDialog()!) {
            Filename = dlg.FileName;
        }
    }

    private static readonly string sqlFrom =
        @"FROM ([Students$] AS Students
        INNER JOIN [Grades$] AS Grades ON Students.CurrentGrade = Grades.CurrentGrade)
        INNER JOIN [Items$] AS Items ON Grades.NewGrade = Items.NewGrade";

    private static readonly (string, string)[] sqlTests = IIFE(() => {
        var ret = new List<(string, string)>();

        // check if sheet exists
        BookStructure.SelectKVP((sheetname, _) => (
            $"SELECT * FROM [{sheetname}$]",
            $"חסר גליון {sheetname} בקובץ Excel"
        )).AddRangeTo(ret);

        // check if field exists on sheet
        BookStructure.Flatten<string, string, List<string>>().SelectT((sheetname, fieldname) => (
            $"SELECT [{fieldname}] FROM [{sheetname}$]",
            $"חסר עמודה {fieldname} בגליון {sheetname}"
        )).AddRangeTo(ret);

        // test the SQL from clause
        ret.Add((
            $"SELECT * {sqlFrom}",
            "שגיאה בצירוף עמודות"
        ));

        // test each field in the SQL from clause
        BookStructure.Flatten<string, string, List<string>>().SelectT((sheetname, fieldname) => (
            $"SELECT [{sheetname}.{fieldname}] {sqlFrom}",
            $"לא מצליח לייבא עמודה {fieldname} מתוך גליון {sheetname}"
        )).AddRangeTo(ret);

        return ret.ToArray();
    });

    private void processFile(object? o) {
        try {
            ProcessState = Success;
            IsRunning = true;

            foreach (var (sql, message) in sqlTests) {
                try {
                    TestSingleSql(sql);
                } catch (Exception ex) {
                    ProcessState = Error;
                    ErrorMessage = ex.Message;
                    return;
                }
            }

            foreach (var dataCheck in DataChecks) {
                var (message, sql, isError) = dataCheck;
                try {
                    var dt = GetDataTable(sql);
                    dataCheck.Data = dt;
                    if (dt.Rows.Count == 0) { continue; }
                    ProcessState = isError ? Warning : Error;
                } catch (Exception ex) {
                    ProcessState = Error;
                    ErrorMessage = ex.Message;
                    return;
                }
            }

            var sqlFinal = $@"
SELECT Students.Name1 & "" "" & Students.Name2, Students.CurrentGrade, Grades.NewGrade, Items.Type, Items.Order, Items.Item, Items.Price, NULL AS Qty
{sqlFrom}
ORDER BY Students.CurrentGrade, Grades.NewGrade, Students.Name1, Students.Name2, Items.Order";

            var rst = GetRst(sqlFinal);
            WriteFinal(rst, Filename);
            ReleaseRst(ref rst);

        } catch (Exception ex) {
            MessageBox.Show(ex.Message);
        } finally {
            IsRunning = false;
        }
    }
}
