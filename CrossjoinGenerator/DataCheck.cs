using System.Data;
using Util;
using static CrossjoinGenerator.ProcessState;

namespace CrossjoinGenerator;
public class DataCheck(string description, string fromClause, bool isError) : ViewModelBase {
    public string Description { get; } = description;
    public string FromClause { get; } = fromClause;

    public ProcessState? State => 
        Data is null ? null :
        RowCount == 0 ? Success :
        isError ? Error :
        Warning;

    private DataTable? data;
    public DataTable? Data {
        get => data;
        set {
            var oldCount = data?.Rows.Count ?? 0;
            var oldState = State;
            NotifyChanged(ref data, value);
            NotifyChanged(ref oldCount, RowCount, nameof(RowCount));
            NotifyChanged(ref oldState, State, nameof(State));
        }
    }

    public int RowCount => Data?.Rows.Count ?? 0;

    public void Deconstruct(out string description, out string fromClause, out bool err) {
        description = Description;
        fromClause = FromClause;
        err = isError;
    }
}
