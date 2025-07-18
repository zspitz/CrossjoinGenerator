using Ookii.Dialogs.Wpf;
using System.Windows;
using Util;
using static CrossjoinGenerator.ExcelFunctions;
using static CrossjoinGenerator.RecordsetFunctions;

namespace CrossjoinGenerator;

public partial class MainWindow : Window {
    public MainWindow() {
        InitializeComponent();

        var vm = new MainViewModel();
        DataContext = vm;

        errorMessage.MouseLeftButtonDown += (s, e) => {
            if (vm.ErrorMessage is null or "" or "...") { return; }
            MessageBox.Show(vm.ErrorMessage, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        };
    }
}
