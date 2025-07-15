using ADODB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Util;

namespace CrossjoinGenerator; 
public class RecordsetFunctions {
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

    public static void TestSingleSql(string sql) {
        var rst = GetRst(sql);
        ReleaseRst(ref rst);
    }
    
    public static void ReleaseRst(ref Recordset rst) {
        if (rst is null) {return;}
        if (rst.State == (int)ObjectStateEnum.adStateOpen) {
            rst.Close();
        }
        Marshal.ReleaseComObject(rst);
        rst = null!;
    }
}
