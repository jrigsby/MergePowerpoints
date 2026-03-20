namespace PdfLibrary;

public interface IPdfHelper {
    List<MemoryStream> SplitPdfStream(MemoryStream inputStream, int pagesPerChunk);
    MemoryStream MergePdfStreams(List<Stream> streams);
    MemoryStream MergePdfFiles(List<string> filePaths);
}