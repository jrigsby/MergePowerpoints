using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;

namespace PdfLibrary.PdfSharp;

// PdfSharpCore goes by MIT License, free for commercial and non-commercial applications

// Dependencies: While PdfSharpCore is MIT-licensed, it relies on SixLabors.ImageSharp
// and SixLabors.Fonts, which are distributed under the Apache 2.0 license when
// used within this package.

// The library is a .NET Core/netstandard 2.0 port of the original PDFsharp. 
public class PdfHelper: IPdfHelper {
    public List<string> SplitPdf(string filePath, int pagesPerChunk, string outputDirectory = "",
        string outputFileNameFormat = "") {
        var files = new List<string>();
        var fileName = Path.GetFileName(filePath);
        var currentDirectory = Directory.GetCurrentDirectory();
        if (string.IsNullOrEmpty(outputDirectory)) {
            outputDirectory = currentDirectory;
        }

        var clonePath = Path.Combine(outputDirectory, fileName);

        if (!File.Exists(clonePath)) {
            File.Copy(filePath,
                clonePath, true);
        }
        
        files.Add(clonePath);
// Open the file
        PdfDocument inputDocument = PdfReader.Open(clonePath, PdfDocumentOpenMode.Import);

        var name = Path.GetFileNameWithoutExtension(fileName);

        var totalPages = inputDocument.PageCount;
        for (var idx = 0; idx < totalPages; idx += pagesPerChunk) {
            {
                // 2. Determine the page range for the current chunk
                var startPage = idx + 1;
                var endPage = Math.Min(idx + pagesPerChunk, totalPages);

                // Create a new document
                PdfDocument outputDocument = new PdfDocument();
                outputDocument.Version = inputDocument.Version;
                outputDocument.Info.Title =
                    $"Page {startPage} of {inputDocument.Info.Title}";
                outputDocument.Info.Creator = inputDocument.Info.Creator;

                // Add the page and save it
                for (var j = startPage; j <= endPage; j += 1) {
                    outputDocument.AddPage(inputDocument.Pages[j - 1]);
                }

                string outputFileName;
                if (!string.IsNullOrEmpty(outputFileNameFormat)) {
                    outputFileName = outputFileNameFormat
                        .Replace("{Name}", name)
                        .Replace("{StartPage}", LeadingZero(startPage))
                        .Replace("{EndPage}", LeadingZero(endPage));
                    if (!outputFileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) {
                        outputFileName += ".pdf";
                    }
                }
                else if (pagesPerChunk == 1) {
                    outputFileName = $"{name} - Page {LeadingZero(startPage)}.pdf";
                }
                else {
                    outputFileName = $"{name} - Pages {LeadingZero(startPage)}_{LeadingZero(endPage)}.pdf";
                }

                var piecePath = Path.Combine(outputDirectory, outputFileName);
                outputDocument.Save(piecePath);
                files.Add(piecePath);
            }
        }

        return files;
    }

    private string LeadingZero(int number) {
        if (number < 10) {
            return "0" + number;
        }

        return number.ToString();
    }

    public List<MemoryStream> SplitPdfStream(MemoryStream inputStream, int pagesPerChunk) {
        var outputStreams = new List<MemoryStream>();

        // 1. Load the document (using your chosen library's specific class, e.g., PdfLoadedDocument)
        using (var document = PdfReader.Open(inputStream, PdfDocumentOpenMode.Import)) {
            int totalPages = document.Pages.Count;

            for (int i = 0; i < totalPages; i += pagesPerChunk) {
                // 2. Determine the page range for the current chunk
                var startPage = i + 1;
                var endPage = Math.Min(i + pagesPerChunk, totalPages);

                // 3. Create a new document with the selected pages
                var newDocument = new PdfDocument();

                // Add the page and save it

                for (int j = startPage; j <= endPage; j += 1) {
                    newDocument.AddPage(document.Pages[j - 1]);
                }

                // 4. Save the new document to a memory stream
                var outputStream = new MemoryStream();
                newDocument.Save(outputStream);

                outputStream.Position = 0; // Reset stream position for reading
                outputStreams.Add(outputStream);
            }

            // 5. Close the original document
            document.Close();
        }

        return outputStreams;
    }

    public MemoryStream MergePdfStreams(List<Stream> streams) {
        var outputStream = new MemoryStream();
        var mergedDocument = new PdfDocument();
        foreach (var stream in streams) {
            using (var document = PdfReader.Open(stream, PdfDocumentOpenMode.Import)) {
                foreach (var page in document.Pages) {
                    mergedDocument.AddPage(page);
                }
            }
        }

        mergedDocument.Save(outputStream);
        mergedDocument.Close();

        outputStream.Position = 0;
        return outputStream;
    }

    public MemoryStream MergePdfFiles(List<string> filePaths) {
        var outputStream = new MemoryStream();
        var mergedDocument = new PdfDocument();
        foreach (var filePath in filePaths) {
            using (var fs = File.OpenRead(filePath)) {
                using (var document = PdfReader.Open(fs, PdfDocumentOpenMode.Import)) {
                    foreach (var page in document.Pages) {
                        mergedDocument.AddPage(page);
                    }
                }
            }
        }

        mergedDocument.Save(outputStream);
        mergedDocument.Close();

        outputStream.Position = 0;
        return outputStream;
    }
}