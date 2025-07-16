using System.Data;
using Util;

namespace CrossjoinGenerator;
public class DataCheck(string description, string fromClause, bool isError) : ViewModelBase {
    public string Description { get; } = description;
    public string FromClause { get; } = fromClause;
    public bool IsError { get; } = isError;

    private DataTable? data;
    public DataTable? Data {
        get => data;
        set {
            var oldCount = data?.Rows.Count ?? 0;
            NotifyChanged(ref data, value);
            NotifyChanged(ref oldCount, RowCount, nameof(RowCount));
        }
    }

    public int RowCount => Data?.Rows.Count ?? 0;

    public void Deconstruct(out string description, out string fromClause, out bool isError) {
        description = Description;
        fromClause = FromClause;
        isError = IsError;
    }
}
