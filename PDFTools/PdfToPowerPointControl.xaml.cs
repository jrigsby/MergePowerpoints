using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Win32;
using PdfLibrary.OpenXml;
using PDFTools.Models;

namespace PDFTools;

public partial class PdfToPowerPointControl : UserControl {
    public ObservableCollection<FileResult> Results { get; set; } = new ObservableCollection<FileResult>();

    public PdfToPowerPointControl() {
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

    private void ConvertPdfToPowerPoint_Click(object sender, RoutedEventArgs e) {
        var pdfPath = PdfFilePathTextBox.Text;
        var outputDirectory = OutputFolderTextBox.Text;
        var filename = FilenameTextBox.Text;

        if (string.IsNullOrEmpty(pdfPath)) {
            MessageBox.Show("Please select a PDF file.");
            return;
        }

        if (string.IsNullOrEmpty(outputDirectory)) {
            MessageBox.Show("Please select an output folder.");
            return;
        }

        if (string.IsNullOrEmpty(filename)) {
            MessageBox.Show("Please enter a filename.");
            return;
        }
        
        try {
            var layout = (LayoutComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
            var templateName = layout switch {
                "Portrait" => "PortraitTemplate.pptx",
                "Landscape" => "LandscapeTemplate.pptx",
                "Square" => "MixTemplate.pptx",
                _ => "PortraitTemplate.pptx"
            };
            
            var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", templateName);
            var destination = Path.Combine(outputDirectory, filename.EndsWith(".pptx") ? filename : filename + ".pptx");

            var pdfSharpHelper = new PdfLibrary.PdfSharp.PdfHelper();
            using (var template = PresentationDocument.CreateFromTemplate(templatePath)) {

                using (var output = template.Clone(destination)) {
                    using (var stream = new MemoryStream(File.ReadAllBytes(pdfPath))) {
                        //PowerPoint conversion is limited to a PDF of 3 pages. To work with this, we will
                        //convert the current PDF into separate files with a max page size of 3.
                        var pieces = pdfSharpHelper.SplitPdfStream(stream, Constants.PdfToPowerPointConversionChunkSize);
                        var i = 0;
                        foreach (var piece in pieces) {
                            i++;
                            Spire.Pdf.PdfDocument doc = new Spire.Pdf.PdfDocument();

                            // 2. Load a sample PDF document
                            doc.LoadFromStream(piece);

                            // 3. Save the document as a PowerPoint document
                            using (MemoryStream temp = new MemoryStream()) {
                                doc.SaveToStream(temp, Spire.Pdf.FileFormat.PPTX);

                                temp.Position = 0;

                                var test = PresentationDocument.Open(temp, true);
                                var powerPointHelper = new PowerPointHelper();
                                var guid = Guid.NewGuid();
                                powerPointHelper.MergePresentationSlidesStreams(test,
                                    guid.ToString()//$"big-{i}.pptx"
                                    , output);
                            }

                            doc.Close();
                        }
                    }

                    var validator = new OpenXmlValidator();

                    var errors = validator.Validate(output);

                    var validationErrorInfos = errors.ToList();
                    if (validationErrorInfos.Any()) {
                        StringBuilder builder = new StringBuilder();
                        builder.AppendLine(
                            "The merge process completed bye the merged presentation failed to validate.");
                        builder.AppendLine("\n");
                        builder.AppendLine("There are " + validationErrorInfos.Count() + " errors:\r\n");
                        builder.AppendLine(PowerPointHelper.FormatErrors(validationErrorInfos));
                    }

                }
            }
            
            Results.Clear();
            
            Results.Add(new FileResult { Path = pdfPath });
            Results.Add(new FileResult { Path = destination });
        }
        catch (Exception ex) {
            MessageBox.Show($"Error converting PDF: {ex.Message}");
        }
    }

    private void OpenFile_Click(object sender, RoutedEventArgs e) {
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