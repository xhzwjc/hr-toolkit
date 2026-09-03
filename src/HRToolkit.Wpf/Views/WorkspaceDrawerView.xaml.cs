using System.Windows;
using System.Windows.Controls;
using HRToolkit.Wpf.ViewModels;

namespace HRToolkit.Wpf.Views
{
    public partial class WorkspaceDrawerView : UserControl
    {
        public WorkspaceDrawerView()
        {
            InitializeComponent();
        }

        private void DropArea_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private async void DropArea_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && DataContext is WorkspaceViewModel vm)
                {
                    await vm.ImportFilesAsync(files);
                }
            }
        }
    }
}
