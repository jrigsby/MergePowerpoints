using DocumentFormat.OpenXml.Packaging;
using PdfLibrary.OpenXml;

namespace MergePowerpoints.OpenXml;

public static class PdfToPptConverter {
    public static void ConvertToSquarePpt(string pdfPath, string outputPath, string templatePath) {
        var powerPointHelper = new PowerPointHelper();
        // Ensure PDFium native binaries are resolvable before first use
        PdfiumLoader.EnsureLoaded();
        var pdfBytes = File.ReadAllBytes(pdfPath);

        using (var pdfDoc = new PDFiumSharp.PdfDocument(pdfBytes)) {
            using (var template = PresentationDocument.CreateFromTemplate(templatePath)) {
                using (var output = template.Clone(outputPath)) {
                    int pageCount = pdfDoc.Pages.Count;

                    // We process pages individually to create a square PPTX via the "working method" (Spire.Pdf)
                    for (int i = 0; i < pageCount; i++) {
                        using (var squareImageStream = PdfImageRenderer.RenderPageToSquarePng(pdfBytes, i)) {
                            // 1. Create a new 11x11 PDF document (square)
                            using (var tempPdf = new Spire.Pdf.PdfDocument()) {
                                // Set page size to 11x11 inches (792x792 points)
                                float sideLength = 11f * 72f;
                                Spire.Pdf.PdfPageBase page = tempPdf.Pages.Add(new SizeF(sideLength, sideLength),
                                    new Spire.Pdf.Graphics.PdfMargins(0));

                                // 2. Draw the square image onto the page
                                var image = Spire.Pdf.Graphics.PdfImage.FromStream(squareImageStream);
                                page.Canvas.DrawImage(image, 0, 0, sideLength, sideLength);

                                // 3. Use the working method: Save to PPTX stream
                                using (var tempPptxStream = new MemoryStream()) {
                                    tempPdf.SaveToStream(tempPptxStream, Spire.Pdf.FileFormat.PPTX);
                                    tempPptxStream.Position = 0;

                                    // 4. Merge into output
                                    using (var sourceDeck = PresentationDocument.Open(tempPptxStream, false)) {
                                        powerPointHelper.MergePresentationSlidesStreams(sourceDeck, $"page-{i}.pptx",
                                            output);
                                    }
                                }
                            }
                        }
                    }

                    output.Save();
                }
            }
        }
    }
}