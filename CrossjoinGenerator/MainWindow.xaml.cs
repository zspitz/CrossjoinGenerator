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
    }
}
