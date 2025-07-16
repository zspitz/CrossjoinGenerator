using System.ComponentModel;
using System.Runtime.CompilerServices;
using static Util.Functions;

namespace Util;
public abstract class ViewModelBase : INotifyPropertyChanged {
    public event PropertyChangedEventHandler? PropertyChanged;

    private void invoke(string? name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    /// <summary>Raises change notification for fields defined in the current class</summary>
    protected void NotifyChanged<T>(ref T current, T newValue, [CallerMemberName] string? name = null) {
        if (IsEqual(current, newValue)) { return; }
        current = newValue;
        invoke(name);
    }

    /// <summary>
    /// Raises change notification for fields not defined in the current class (e.g. the model class)
    /// because fields defined in another class cannot be passed with the ref modifier
    /// </summary>
    protected void NotifyChanged<T>(T current, T newValue, Action? setter = null, [CallerMemberName] string? name = null) {
        if (IsEqual(current, newValue)) { return; }
        setter?.Invoke();
        invoke(name);
    }

    ///<summary>Raise change notification without checking for equality first.</summary>
    protected void NotifyChanged([CallerMemberName] string? name = null) => invoke(name);
}
