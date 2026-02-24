using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using PDFTools.Models;

namespace PDFTools;

public partial class PdfSplitControl : UserControl {
    public ObservableCollection<FileResult> Results { get; set; } = new ObservableCollection<FileResult>();

    public PdfSplitControl() {
        InitializeComponent();
        ResultsGrid.ItemsSource = Results;
    }

    private void BrowseFile_Click(object sender, RoutedEventArgs e) {
        var openFileDialog = new OpenFileDialog {
            Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
        };
        if (openFileDialog.ShowDialog() == true) {
            PdfFilePathTextBox.Text = openFileDialog.FileName;
        }
    }

    private void BrowseFolder_Click(object sender, RoutedEventArgs e) {
        var openFolderDialog = new OpenFolderDialog();
        if (openFolderDialog.ShowDialog() == true) {
            OutputFolderTextBox.Text = openFolderDialog.FolderName;
        }
    }

    private void SplitPdf_Click(object sender, RoutedEventArgs e) {
        var filePath = PdfFilePathTextBox.Text;
        var outputDirectory = OutputFolderTextBox.Text;
        var filenameFormat = FilenameFormatTextBox.Text;

        if (string.IsNullOrEmpty(filePath)) {
            MessageBox.Show("Please select a PDF file.");
            return;
        }

        if (!int.TryParse(PagesPerSplitTextBox.Text, out var pagesPerChunk)) {
            MessageBox.Show("Please enter a valid number for pages per split.");
            return;
        }

        try {
            //PdfLibrary.PdfSharp.PdfHelper.SplitPdf(pdfPath, 1, "", "{Name} - Page {StartPage}.pdf");
            //PdfLibrary.PdfSharp.PdfHelper.SplitPdf(pdfPath, Constants.ChunkSize);
            // using (var stream = new MemoryStream(File.ReadAllBytes(pdfPath))) {
            //     var x = PdfLibrary.PdfSharp.PdfHelper.SplitPdfStream(stream, Constants.ChunkSize);
            // }
            var files = PdfLibrary.PdfSharp.PdfHelper.SplitPdf(filePath, pagesPerChunk, outputDirectory, filenameFormat);
            Results.Clear();
            foreach (var file in files) {
                Results.Add(new FileResult { Path = file });
            }
        }
        catch (Exception ex) {
            MessageBox.Show($"Error splitting PDF: {ex.Message}");
        }
    }

    private void OpenPdf_Click(object sender, RoutedEventArgs e) {
        if (sender is Button button && button.DataContext is FileResult result) {
            try {
                Process.Start(new ProcessStartInfo {
                    FileName = result.Path,
                    UseShellExecute = true
                });
            }
            catch (Exception ex) {
                MessageBox.Show($"Error opening file: {ex.Message}");
            }
        }
    }
}
