using Ookii.Dialogs.Wpf;
using System.Data;
using System.Windows;
using Util;
using static CrossjoinGenerator.ExcelFunctions;
using static CrossjoinGenerator.ProcessState;
using static CrossjoinGenerator.RecordsetFunctions;
using static Util.Functions;

namespace CrossjoinGenerator;

public class MainViewModel : ViewModelBase {
    public RelayCommand GenerateTemplate { get; } = new(o => GenerateTemplate());

    public RelayCommand ChooseFile { get; }
    public RelayCommand ProcessFile { get; }
    public RelayCommand EditFile { get; }

    private string filename = "";
    public string Filename {
        get => filename;
        set {
            NotifyChanged(ref filename, value);
            RecordsetFunctions.Filename = value;
        }
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

    private string progressCaption = "";
    public string ProgressCaption {
        get => progressCaption;
        set => NotifyChanged(ref progressCaption, value);
    }

    private double progressValue = 0;
    public double ProgressValue {
        get => progressValue;
        set => NotifyChanged(ref progressValue, value);
    }

    private string errorMessage = "...";
    public string ErrorMessage {
        get => errorMessage;
        set => NotifyChanged(ref errorMessage, value);
    }

    public double MaxProgess => sqlTests.Length + DataChecks.Count + 2;

    public List<DataCheck> DataChecks { get; } = [
        new(
            "תלמיד ללא שם",
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
            false
        ),
        new(
            "כיתה ללא תלמידים",
            @"
                SELECT *
                FROM [Grades$] AS Grades
                LEFT JOIN [Students$] AS Students ON Grades.CurrentGrade = Students.CurrentGrade
                WHERE Students.CurrentGrade IS NULL
            ",
            false
        ),
        new(
            "כתה ללא פריטים",
            @"
                SELECT *
                FROM [Grades$] AS Grades
                LEFT JOIN [Items$] AS Items ON Grades.NewGrade = Items.NewGrade
                WHERE Items.NewGrade IS NULL
            ",
            false
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

    private static readonly (string, string, string)[] sqlTests = IIFE(() => {
        var ret = new List<(string, string, string)>();

        // check if sheet exists
        BookStructure.SelectKVP((sheetname, _) => (
            $"SELECT * FROM [{sheetname}$]",
            $"חסר גליון {sheetname} בקובץ Excel",
            $"בדיקת קיום גליון ${sheetname}"
        )).AddRangeTo(ret);

        // check if field exists on sheet
        BookStructure.Flatten<string, string, List<string>>().SelectT((sheetname, fieldname) => (
            $"SELECT [{fieldname}] FROM [{sheetname}$]",
            $"חסר עמודה {fieldname} בגליון {sheetname}",
            $"בדיקת קיום עמודה {fieldname} בגליון ${sheetname}"
        )).AddRangeTo(ret);

        // test the SQL from clause
        ret.Add((
            $"SELECT * {sqlFrom}",
            "שגיאה בצירוף גליונות",
            "בדיקת צירוף גליונות"
        ));

        // test each field in the SQL from clause
        BookStructure.Flatten<string, string, List<string>>().SelectT((sheetname, fieldname) => (
            $"SELECT [{sheetname}.{fieldname}] {sqlFrom}",
            $"לא מצליח לייבא עמודה {fieldname} מתוך גליון {sheetname}",
            $"בדיקת ייבוא עמודה {fieldname} מתוך גליון {sheetname}"
        )).AddRangeTo(ret);

        return ret.ToArray();
    });

    private async void processFile(object? o) {
        try {
            ProcessState = null;
            ErrorMessage = "";
            ProgressValue = 0;
            IsRunning = true;

            foreach (var dc in DataChecks) {
                dc.Data = new System.Data.DataTable();
            }

            ProgressCaption = "בדיקות מבנה קובץ (1/3) ...";
            foreach (var (sql, message, caption) in sqlTests) {
                try {
                    await TestSingleSql(sql);
                } catch {
                    ProcessState = Error;
                    ErrorMessage = message;
                    return;
                }
                ProgressValue += 1;
            }

            ProgressCaption = "בדיקות תקינות נתונים (2/3)";
            foreach (var dataCheck in DataChecks) {
                var (description, sql, isError) = dataCheck;
                try {
                    var dt = await GetDataTable(sql);
                    dataCheck.Data = dt;
                    if (dt.Rows.Count != 0) {
                        if (isError) {
                            ProcessState = Error;
                            ErrorMessage = description;
                        } else {
                            ProcessState = Warning;
                        }
                    }
                } catch (Exception ex) {
                    ProcessState = Error;
                    ErrorMessage = ex.Message;
                    return;
                }
                ProgressValue += 1;
            }

            if (
                ProcessState == Error || 
                (
                    ProcessState == Warning && 
                    MessageBoxResult.No == MessageBox.Show("ישנם התראות. האם להמשיך?", "", MessageBoxButton.YesNo)
                )
            ) { return; }

            var sqlFinal = $@"
SELECT Students.Name1 & "" "" & Students.Name2, Students.CurrentGrade, Grades.NewGrade, Items.Type, Items.Order, Items.Item, Items.Price, NULL AS Qty
{sqlFrom}
ORDER BY Students.CurrentGrade, Grades.NewGrade, Students.Name1, Students.Name2, Items.Order";

            ProgressCaption = "בניית טבלא סופית (3/3) ...";

            var dtFinal = await Task.Run(() => {
                var rst = GetRst(sqlFinal);
                var dt = rst.ToDataTable();
                ReleaseRst(ref rst);
                return dt;
            });
            ProgressValue += 1;

            await Task.Run(() => {
                WriteFinal(dtFinal, Filename);
            });
            ProgressValue += 1;
            ProgressCaption = "הקובץ מוכן!";

        } catch (Exception ex) {
            MessageBox.Show(ex.Message);
        } finally {
            IsRunning = false;
        }
    }
}
