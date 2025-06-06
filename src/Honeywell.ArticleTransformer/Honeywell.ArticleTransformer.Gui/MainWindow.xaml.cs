using Microsoft.Win32;
using System.IO;
using System.Windows;
using Honeywell.ArticleTransformer.Core;

namespace Honeywell.ArticleTransformer.Gui
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnSelectFile(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "CSV or XLSX|*.csv;*.xlsx"
            };
            if (dialog.ShowDialog() == true)
            {
                FilePathBox.Text = dialog.FileName;
            }
        }

        private void OnStart(object sender, RoutedEventArgs e)
        {
            var path = FilePathBox.Text;
            if (!File.Exists(path))
            {
                MessageBox.Show("Please select a valid file.");
                return;
            }

            var outputExtension = Path.GetExtension(path).Equals(".xlsx", System.StringComparison.OrdinalIgnoreCase)
                ? ".xlsx" : ".csv";
            var outputPath = Path.Combine(Path.GetDirectoryName(path) ?? string.Empty, $"output{outputExtension}");
            Transformer.ProcessFile(path, outputPath);
            MessageBox.Show($"Processed file saved to {outputPath}");
        }
    }
}
