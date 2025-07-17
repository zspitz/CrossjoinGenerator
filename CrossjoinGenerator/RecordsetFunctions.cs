using ADODB;
using System.Data;
using System.Runtime.InteropServices;
using Util;

namespace CrossjoinGenerator;
public static class RecordsetFunctions {
    public static string Filename { get; set; } = "";

    public static Recordset GetRst(string sql) {
        var connString = new[] {
            @"Provider=""Microsoft.ACE.OLEDB.12.0""",
            @$"Data Source=""{Filename}""",
            @"Extended Properties=""Excel 12.0;HDR=Yes"""
        }.Joined(";");
        var rst = new Recordset();
        rst.Open(Source: sql, ActiveConnection: connString, CursorType: CursorTypeEnum.adOpenForwardOnly, LockType: LockTypeEnum.adLockReadOnly);
        return rst;
    }

    public static Task TestSingleSql(string sql) =>
        Task.Run(() => {
            var rst = GetRst(sql);
            ReleaseRst(ref rst);
        });

    public static void ReleaseRst(ref Recordset rst) {
        if (rst is null) { return; }
        if (rst.State == (int)ObjectStateEnum.adStateOpen) {
            rst.Close();
        }
        Marshal.ReleaseComObject(rst);
        rst = null!;
    }

    public static Task<DataTable> GetDataTable(string sql) =>
        Task.Run(() => {
            var dt = new DataTable();
            var rs = GetRst(sql);

            for (var i = 0; i < rs.Fields.Count; i++) {
                var field = rs.Fields[i];
                dt.Columns.Add(field.Name.Replace('.', '~'), typeof(object));
            }

            // Add rows
            while (!rs.EOF) {
                var row = dt.NewRow();
                for (var i = 0; i < rs.Fields.Count; i++) {
                    row[i] = rs.Fields[i].Value;
                }

                dt.Rows.Add(row);
                rs.MoveNext();
            }

            ReleaseRst(ref rs);

            return dt;
        });
}
