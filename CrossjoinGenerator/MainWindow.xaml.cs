using System.Windows;

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

        helpButton.Click += (s, e) => {
            var help = new HelpWindow();
            help.ShowDialog();
        };
    }

}
