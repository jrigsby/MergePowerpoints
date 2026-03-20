using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using PDFTools.Models;

namespace PDFTools;

public partial class MergePdfsControl : UserControl, INotifyPropertyChanged {
    public ObservableCollection<FileResult> Results { get; set; } = new ObservableCollection<FileResult>();

    private ObservableCollection<ToastViewModel> _toasts = new();
    
    public ObservableCollection<ToastViewModel> Toasts {
        get => _toasts;
        set => SetProperty(ref _toasts, value);
    }
    
    public MergePdfsControl() {
        InitializeComponent();
        ResultsGrid.ItemsSource = Results;
    }

    private void BrowseFile_Click(object sender, RoutedEventArgs e) {
        var openFileDialog = new OpenFileDialog {
            Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*",
            Multiselect = true
        };
        if (openFileDialog.ShowDialog() == true) {
            Results.Clear();
            foreach (var file in openFileDialog.FileNames) {
                Results.Add(new FileResult { Path = file });
            }
        }
    }

    private void BrowseFolder_Click(object sender, RoutedEventArgs e) {
        var openFolderDialog = new OpenFolderDialog();
        if (openFolderDialog.ShowDialog() == true) {
            OutputFolderTextBox.Text = openFolderDialog.FolderName;
        }
    }

    private void MergePdfs_Click(object sender, RoutedEventArgs e) {
        var outputDirectory = OutputFolderTextBox.Text;
        var filename = OutputFilenameTextBox.Text;

        if (string.IsNullOrEmpty(outputDirectory)) {
            MessageBox.Show("Please select a output directory.");
            return;
        }
        
        if (string.IsNullOrEmpty(filename)) {
            MessageBox.Show("Please select a file name.");
            return;
        }
        
        try {
            var pdfSharpHelper = new PdfLibrary.PdfSharp.PdfHelper();
            
            var path = Path.Combine(outputDirectory, filename);
            var findOutput = Results.SingleOrDefault(x => x.Path==path);
            if(findOutput!=null) Results.Remove(findOutput);
            if(File.Exists(path)) File.Delete(path);
            
            //option 1
            // var s = new List<Stream>();
            // foreach (var file in Results.Select(x => x.Path).ToList()) {
            //     s.Add(new MemoryStream(File.ReadAllBytes(file)));
            // }
            // var stream = pdfSharpHelper.MergePdfStreams(s);
            //option 2
            var stream = pdfSharpHelper.MergePdfFiles(Results.Select(x=>x.Path).ToList());
            
            FileStream targetStream = File.OpenWrite(path);
            byte[] buffer = new byte[2048];
            while (true) {
                var bytesRead = stream.Read(buffer, 0, buffer.Length);
                if (bytesRead == 0)
                    break;
                targetStream.Write(buffer, 0, bytesRead);
            }
            stream.Close();
            targetStream.Close();
            Results.Add(new FileResult { Path = path });

            ShowToast($"PDF files merged. File {path} added to list.");
        }
        catch (Exception ex) {
            MessageBox.Show($"Error merging PDF: {ex.Message}");
        }
    }
    
    private void SyncfusionMergePdfs_Click(object sender, RoutedEventArgs e) {
        var outputDirectory = OutputFolderTextBox.Text;
        var filename = OutputFilenameTextBox.Text;

        if (string.IsNullOrEmpty(outputDirectory)) {
            MessageBox.Show("Please select a output directory.");
            return;
        }
        
        if (string.IsNullOrEmpty(filename)) {
            MessageBox.Show("Please select a file name.");
            return;
        }
        
        try {
            var syncFusionHelper = new PdfLibrary.Syncfusion.PdfHelper();
            var path = Path.Combine(outputDirectory, filename);
            var findOutput = Results.SingleOrDefault(x => x.Path==path);
            if(findOutput!=null) Results.Remove(findOutput);
            if(File.Exists(path)) File.Delete(path);
            
            //option 1
            // var s = new List<FileStream>();
            // foreach (var file in Results.Select(x => x.Path).ToList()) {
            //     s.Add(new FileStream(file, FileMode.Open, FileAccess.Read));
            // }
            // var stream = syncFusionHelper.MergePdfStreams(s);
            
            //option 2
            // var s = new List<Stream>();
            // foreach (var file in Results.Select(x => x.Path).ToList()) {
            //     s.Add(new MemoryStream(File.ReadAllBytes(file)));
            // }
            // var stream = syncFusionHelper.MergePdfStreams(s);
            
            //option 3
            var stream = syncFusionHelper.MergePdfFiles(Results.Select(x=>x.Path).ToList());
            
            FileStream targetStream = File.OpenWrite(path);
            byte[] buffer = new byte[2048];
            while (true) {
                var bytesRead = stream.Read(buffer, 0, buffer.Length);
                if (bytesRead == 0)
                    break;
                targetStream.Write(buffer, 0, bytesRead);
            }
            stream.Close();
            targetStream.Close();
            Results.Add(new FileResult { Path = path });

            ShowToast($"PDF files merged. File {path} added to list.");
        }
        catch (Exception ex) {
            MessageBox.Show($"Error merging PDF: {ex.Message}");
        }
    }
    
    private void TelerikMergePdfs_Click(object sender, RoutedEventArgs e) {
        var outputDirectory = OutputFolderTextBox.Text;
        var filename = OutputFilenameTextBox.Text;

        if (string.IsNullOrEmpty(outputDirectory)) {
            MessageBox.Show("Please select a output directory.");
            return;
        }
        
        if (string.IsNullOrEmpty(filename)) {
            MessageBox.Show("Please select a file name.");
            return;
        }
        
        try {
            var telerikHelper = new PdfLibrary.Telerik.TelerikPdfHelper();
            var path = Path.Combine(outputDirectory, filename);
            var findOutput = Results.SingleOrDefault(x => x.Path==path);
            if(findOutput!=null) Results.Remove(findOutput);
            if(File.Exists(path)) File.Delete(path);
            
            //option 1
            // var s = new List<Stream>();
            // foreach (var file in Results.Select(x => x.Path).ToList()) {
            //     s.Add(new MemoryStream(File.ReadAllBytes(file)));
            // }
            // var stream = telerikHelper.MergePdfStreams(s);
            
            //option2
            var stream = telerikHelper.MergePdfFiles(Results.Select(x=>x.Path).ToList());
            
            FileStream targetStream = File.OpenWrite(path);
            byte[] buffer = new byte[2048];
            while (true) {
                var bytesRead = stream.Read(buffer, 0, buffer.Length);
                if (bytesRead == 0)
                    break;
                targetStream.Write(buffer, 0, bytesRead);
            }
            stream.Close();
            targetStream.Close();
            Results.Add(new FileResult { Path = path });

            ShowToast($"PDF files merged. File {path} added to list.");
        }
        catch (Exception ex) {
            MessageBox.Show($"Error merging PDF: {ex.Message}");
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

    private void ShowToast(string message) {
        Application.Current.Dispatcher.Invoke(() => {
            // Avoid duplicate toasts with same message
            if (Toasts.Any(t => t.Message == message)) return;
            
            var toast = new ToastViewModel(message, t => {
                Application.Current.Dispatcher.Invoke(() => Toasts.Remove(t));
            });
            Toasts.Add(toast);
        });
    }
    
    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null) {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string? propertyName = null)
    {
        if (Equals(storage, value)) return false;
        storage = value;
        OnPropertyChanged(propertyName);
        return true;
    }

}