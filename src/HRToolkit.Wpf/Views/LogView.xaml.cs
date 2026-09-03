using System.Windows;
using System.Windows.Controls;
using HRToolkit.Wpf.ViewModels;

namespace HRToolkit.Wpf.Views
{
    public partial class LogView : UserControl
    {
        public LogView()
        {
            InitializeComponent();
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is LogViewModel vm)
            {
                vm.Clear();
            }
        }
    }
}
