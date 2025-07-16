using System.Windows.Input;

namespace Util;

public class RelayCommand(Action<object?> execute, Predicate<object?>? canExecute = null) : ICommand {
    private readonly Action<object?> execute = execute ?? throw new ArgumentNullException(nameof(execute));

    public bool CanExecute(object? parameter) => canExecute == null || canExecute(parameter);

    public event EventHandler? CanExecuteChanged {
        add { CommandManager.RequerySuggested += value; }
        remove { CommandManager.RequerySuggested -= value; }
    }

    public void Execute(object? parameter) => execute(parameter);
}
