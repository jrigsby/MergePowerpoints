using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MergePowerpoints.OpenXml;
using PdfLibrary.OpenXml;

namespace MergePowerpoints;
//https://www.ericwhite.com/blog/merging-powerpoint-presentations/
//https://www.e-iceblue.com/Introduce/free-pdf-component.html#.WZ0VBnS-sl0
class Program {
    //private static uint uniqueId;
    static void Main(string[] args)
    {
        // if (args.Length < 3)
        // {
        //     Console.WriteLine("Usage: MergePowerpoints <presentation1.pptx> <presentation2.pptx> <output.pptx>");
        //     return;
        // }

        var powerPointHelper = new PowerPointHelper();
        
        //List<string> files = [];

        //string templatePath ="C:\\Temp\\LandscapeTemplate.pptx";
        //string templatePath ="D:\\Temp\\MixTemplate.pptx";
        string destination ="D:\\Temp\\merged.pptx";
        string pdfPath = "D:\\Temp\\big.pdf";
        try {
            //PdfToPptConverter.ConvertToSquarePpt(pdfPath, "D:\\Temp\\merged2.pptx", "D:\\Temp\\MixTemplate.pptx");
            //using (var template = PresentationDocument.CreateFromTemplate(templatePath)) {

           // using (var template = PresentationDocumentHelper.CreatePresenationAsMemoryStreamFromPdf(pdfPath)) {
                //using (PresentationDocument template = PresentationDocument.Open(memStream, isEditable: true)){
                using (PresentationDocument template = PresentationDocumentHelper.CreateFromPdf(pdfPath)) {

                using (var output = template.Clone(destination)) {
                    // 3. Perform modifications on the new document (newDoc).
                    // Example: add slides, change content, etc.

                    // 4. All changes are automatically saved when the new document is disposed of (at the end of the using statement).

                    var telerikHelper = new PdfLibrary.Telerik.TelerikPdfHelper();
                    using (var stream = new MemoryStream(File.ReadAllBytes(pdfPath))) {
                        var pieces = telerikHelper.SplitPdfStream(stream, Constants.ChunkSize);
                        int i = 0;
                        foreach (var piece in pieces) {
                            i++;
                            Spire.Pdf.PdfDocument doc = new Spire.Pdf.PdfDocument();

// 2. Load a sample PDF document
                            doc.LoadFromStream(piece);

// 3. Save the document as a PowerPoint document
                            // files.Add($"C:\\Temp\\big-{i}.pptx");
                            using (MemoryStream temp = new MemoryStream()) {
                                doc.SaveToStream(temp, Spire.Pdf.FileFormat.PPTX);

                                temp.Position = 0;

                                var test = PresentationDocument.Open(temp, true);

                                powerPointHelper.MergePresentationSlidesStreams(test,
                                    $"big-{i}.pptx"
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

            Console.WriteLine($"Successfully merged files into {destination}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error merging presentations: {ex.Message}");
        }
    }
}