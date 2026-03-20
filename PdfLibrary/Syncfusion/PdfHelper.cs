using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;

namespace PdfLibrary.Syncfusion;

public class PdfHelper: IPdfHelper {
    public List<MemoryStream> SplitPdfStream(MemoryStream inputStream, int pagesPerChunk) {
        var outputStreams = new List<MemoryStream>();

        // 1. Load the document (using your chosen library's specific class, e.g., PdfLoadedDocument)
        using (var document = new PdfLoadedDocument(inputStream)) {
            int totalPages = document.Pages.Count;

            for (int i = 0; i < totalPages; i += pagesPerChunk) {
                // 2. Determine the page range for the current chunk
                var startPage = i + 1;
                var endPage = Math.Min(i + pagesPerChunk, totalPages);

                // 3. Create a new document with the selected pages
                var newDocument = new PdfDocument();

                newDocument.ImportPageRange(document, startPage - 1, endPage - 1);

                // 4. Save the new document to a memory stream
                var outputStream = new MemoryStream();
                newDocument.Save(outputStream);

                outputStream.Position = 0; // Reset stream position for reading
                outputStreams.Add(outputStream);
            }

            // 5. Close the original document
            document.Close(true);
        }

        return outputStreams;
    }

    public MemoryStream MergePdfStreams(List<Stream> streams) {
        var outputStream = new MemoryStream();
        var mergedDocument = new PdfDocument();


        PdfDocumentBase.Merge(mergedDocument, streams.ToArray());

        mergedDocument.Save(outputStream);
        mergedDocument.Close(true);

        outputStream.Position = 0;
        return outputStream;
    }

    public MemoryStream MergePdfStreams(List<FileStream> fileStreams) {
        var outputStream = new MemoryStream();
        var mergedDocument = new PdfDocument();

        PdfDocumentBase.Merge(mergedDocument, fileStreams.ToArray<Stream>());

        mergedDocument.Save(outputStream);
        mergedDocument.Close(true);

        outputStream.Position = 0;
        return outputStream;
    }

    public MemoryStream MergePdfFiles(List<string> filePaths) {
        var outputStream = new MemoryStream();
        var mergedDocument = new PdfDocument();
        foreach (var filePath in filePaths) {
            using (FileStream stream1 = new FileStream(filePath, FileMode.Open, FileAccess.Read)) {
                Stream[] streams = { stream1 };
                PdfDocumentBase.Merge(mergedDocument, streams);
            }
        }
        
        mergedDocument.Save(outputStream);
        mergedDocument.Close(true);

        outputStream.Position = 0;
        return outputStream;
    }
}