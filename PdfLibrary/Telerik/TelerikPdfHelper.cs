using Telerik.Windows.Documents.Fixed.FormatProviders.Pdf;
using Telerik.Windows.Documents.Fixed.FormatProviders.Pdf.Streaming;
using Telerik.Windows.Documents.Fixed.Model;

namespace PdfLibrary.Telerik;

public static class TelerikPdfHelper {
    public static List<MemoryStream> SplitPdfStream(MemoryStream inputStream, int pagesPerChunk) {
        var outputStreams = new List<MemoryStream>();
        PdfFormatProvider provider = new PdfFormatProvider();
        RadFixedDocument document = provider.Import(inputStream);

        int totalPages = document.Pages.Count;

        for (int i = 0; i < totalPages; i += pagesPerChunk) {
            RadFixedDocument newDocument = new RadFixedDocument();
            var endPage = Math.Min(i + pagesPerChunk, totalPages);

            for (int pageIndex = i; pageIndex < endPage; pageIndex++) {
                RadFixedPage page = document.Pages[pageIndex];

                // Temporarily move the page to a new document to export it (cloning)
                RadFixedDocument tempDoc = new RadFixedDocument();
                document.Pages.RemoveAt(pageIndex);
                tempDoc.Pages.Add(page);

                using (MemoryStream ms = new MemoryStream()) {
                    provider.Export(tempDoc, ms);
                    ms.Position = 0;

                    // Restore original document structure
                    tempDoc.Pages.Remove(page);
                    document.Pages.Insert(pageIndex, page);

                    // Import the exported page into the chunk document
                    RadFixedDocument clonedDoc = provider.Import(ms);
                    RadFixedPage clonedPage = clonedDoc.Pages[0];
                    clonedDoc.Pages.Remove(clonedPage);
                    newDocument.Pages.Add(clonedPage);
                }
            }

            var outputStream = new MemoryStream();
            provider.Export(newDocument, outputStream);
            outputStream.Position = 0;
            outputStreams.Add(outputStream);
        }

        return outputStreams;
    }

    public static MemoryStream MergePdfStreams(List<Stream> streams) {
        var outputStream = new MemoryStream();
        using (PdfStreamWriter fileWriter = new PdfStreamWriter(outputStream, leaveStreamOpen: true))
        {
            foreach (var sourceStream in streams)
            {
                using (PdfFileSource fileSource = new PdfFileSource(sourceStream))
                {
                    foreach (var page in fileSource.Pages)
                    {
                        using (var pageWriter = fileWriter.BeginPage(page.Size))
                        {
                            pageWriter.WriteContent(page);
                        }
                    }
                }
            }
        }
        outputStream.Position = 0;
        return outputStream;
    }

    public static MemoryStream MergePDFFiles(List<string> filePaths) {
        var outputStream = new MemoryStream();
        using (PdfStreamWriter fileWriter = new PdfStreamWriter(outputStream, leaveStreamOpen: true))
        {
            foreach (var filePath in filePaths)
            {
                using (PdfFileSource fileSource = new PdfFileSource(File.OpenRead(filePath)))
                {
                    foreach (var page in fileSource.Pages)
                    {
                        using (var pageWriter = fileWriter.BeginPage(page.Size))
                        {
                            pageWriter.WriteContent(page);
                        }
                    }
                }
            }
        }
        outputStream.Position = 0;
        return outputStream;
        }
    }