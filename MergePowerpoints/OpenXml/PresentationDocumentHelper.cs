using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace MergePowerpoints.OpenXml;

// Builds a minimal valid PPTX in memory (as a ZIP) sized to match the first page of a PDF.
// Avoids Spire.Presentation.SaveToFile which requires a compatible SkiaSharp version.
public static class PresentationDocumentHelper {
    private const long PointsToEmu = 12700L; // 1 pt = 12700 EMU; Spire.Pdf reports in points

    public static PresentationDocument CreateFromPdf(string pdfPath) {
        var pdfDoc = new Spire.Pdf.PdfDocument();
        pdfDoc.LoadFromFile(pdfPath);
        var pageSize = pdfDoc.Pages[0].Size;
        pdfDoc.Close();

        long widthEmu = (long)(pageSize.Width * PointsToEmu);
        long heightEmu = (long)(pageSize.Height * PointsToEmu);
        return BuildPresentation(widthEmu, heightEmu);
    }

    public static PresentationDocument CreateFromPdf(MemoryStream pdfStream) {
        var pdfDoc = new Spire.Pdf.PdfDocument();
        pdfDoc.LoadFromStream(pdfStream);
        var pageSize = pdfDoc.Pages[0].Size;
        pdfDoc.Close();

        long widthEmu = (long)(pageSize.Width * PointsToEmu);
        long heightEmu = (long)(pageSize.Height * PointsToEmu);
        return BuildPresentation(widthEmu, heightEmu);
    }

    // public static void CreatePresentation(string filepath) {
    //     using (PresentationDocument presentationDoc =
    //            PresentationDocument.Create(filepath, PresentationDocumentType.Presentation)) {
    //         PresentationPart presentationPart = presentationDoc.AddPresentationPart();
    //         presentationPart.Presentation = new Presentation();
    //
    //         // Slide Size (13.33 x 7.5 inches in EMUs)
    //         SlideSize slideSize = new SlideSize() {
    //             Cx = 12192000,
    //             Cy = 6858000,
    //             Type = SlideSizeValues.Screen16x9
    //         };
    //
    //         NotesSize notesSize = new NotesSize() { Cx = 6858000, Cy = 9144000 };
    //         presentationPart.Presentation.Append(slideSize, notesSize);
    //
    //         // Create Slide Master
    //         SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
    //
    //         // FIX: ColorMap properties corrected to Background1, Text1, etc.
    //         ColorMap colorMap = new ColorMap() {
    //             Background1 = A.ColorSchemeIndexValues.Light1,
    //             Text1 = A.ColorSchemeIndexValues.Dark1,
    //             Background2 = A.ColorSchemeIndexValues.Light2,
    //             Text2 = A.ColorSchemeIndexValues.Dark2,
    //             Accent1 = A.ColorSchemeIndexValues.Accent1,
    //             Accent2 = A.ColorSchemeIndexValues.Accent2,
    //             Accent3 = A.ColorSchemeIndexValues.Accent3,
    //             Accent4 = A.ColorSchemeIndexValues.Accent4,
    //             Accent5 = A.ColorSchemeIndexValues.Accent5,
    //             Accent6 = A.ColorSchemeIndexValues.Accent6,
    //             Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
    //             FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
    //         };
    //
    //         slideMasterPart.SlideMaster = new SlideMaster(
    //             new CommonSlideData(new ShapeTree(new NonVisualGroupShapeProperties(), new GroupShapeProperties(),
    //                 new A.TransformGroup())),
    //             colorMap,
    //             new SlideLayoutIdList()
    //         );
    //
    //         // Create Slide Layout
    //         SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
    //         slideLayoutPart.SlideLayout = new SlideLayout(
    //             new CommonSlideData(new ShapeTree(new NonVisualGroupShapeProperties(), new GroupShapeProperties(),
    //                 new A.TransformGroup()))
    //         );
    //
    //         slideMasterPart.SlideMaster.SlideLayoutIdList.Append(new SlideLayoutId()
    //             { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) });
    //
    //         // Create the Slide
    //         SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
    //         slidePart.Slide = new Slide(
    //             new CommonSlideData(
    //                 new ShapeTree(
    //                     new NonVisualGroupShapeProperties(
    //                         new NonVisualDrawingProperties() { Id = 1U, Name = "" },
    //                         new NonVisualGroupShapeDrawingProperties(),
    //                         new ApplicationNonVisualDrawingProperties()
    //                     ),
    //                     new GroupShapeProperties(new A.TransformGroup()),
    //                     new Shape(
    //                         new NonVisualShapeProperties(
    //                             new NonVisualDrawingProperties() { Id = 2U, Name = "Background" },
    //                             new NonVisualShapeDrawingProperties(),
    //                             new ApplicationNonVisualDrawingProperties()
    //                         ),
    //                         new ShapeProperties(
    //                             new A.Transform2D(
    //                                 new A.Offset() { X = 0L, Y = 0L },
    //                                 new A.Extents() { Cx = 12192000, Cy = 6858000 }
    //                             ),
    //                             // FIX: Changed Prst to Preset
    //                             new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle }
    //                         )
    //                     )
    //                 )
    //             )
    //         );
    //
    //         slidePart.AddPart(slideLayoutPart);
    //
    //         // Finalize structure
    //         SlideIdList slideIdList = new SlideIdList();
    //         SlideId slideId = new SlideId() {
    //             Id = 256U,
    //             RelationshipId = presentationPart.GetIdOfPart(slidePart)
    //         };
    //         slideIdList.Append(slideId);
    //         presentationPart.Presentation.Append(slideIdList);
    //
    //         presentationPart.Presentation.Save();
    //     }
    // }

    // public static MemoryStream CreatePresenationAsMemoryStreamFromPdf(string pdfPath) {
    //     var pdfDoc = new Spire.Pdf.PdfDocument();
    //     pdfDoc.LoadFromFile(pdfPath);
    //     var pageSize = pdfDoc.Pages[0].Size;
    //     pdfDoc.Close();
    //
    //     long widthEmu = (long)(pageSize.Width * PointsToEmu);
    //     long heightEmu = (long)(pageSize.Height * PointsToEmu);
    //     return CreateInMemoryBlankTemplate(widthEmu, heightEmu);
    // }
    //
    // public static MemoryStream CreateInMemoryBlankTemplate(long emuWidth, long emuHeight)
    // {
    //     MemoryStream stream = new MemoryStream();
    //
    //     using (PresentationDocument presentationDoc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
    //     {
    //         // 1. Initialize Root Presentation Part
    //         PresentationPart presentationPart = presentationDoc.AddPresentationPart();
    //         presentationPart.Presentation = new Presentation();
    //
    //         // 2. Setup Parts Tree
    //         SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
    //         SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
    //         ThemePart themePart = slideMasterPart.AddNewPart<ThemePart>();
    //
    //         // Bidirectional tree connection
    //         slideLayoutPart.AddPart(slideMasterPart);
    //
    //         // 3. Populate Theme Data (Fixes #1 and #2: Min counts met, ComplexScript added)
    //         themePart.Theme = new A.Theme() { Name = "Office Theme" };
    //         themePart.Theme.Append(
    //             new A.ThemeElements(
    //                 new A.ColorScheme(
    //                     new A.Dark1Color(new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
    //                     new A.Light1Color(new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
    //                     new A.Dark2Color(new A.RgbColorModelHex() { Val = "1F497D" }),
    //                     new A.Light2Color(new A.RgbColorModelHex() { Val = "EEECE1" }),
    //                     new A.Accent1Color(new A.RgbColorModelHex() { Val = "4F81BD" }),
    //                     new A.Accent2Color(new A.RgbColorModelHex() { Val = "C0504D" }),
    //                     new A.Accent3Color(new A.RgbColorModelHex() { Val = "9BBB59" }),
    //                     new A.Accent4Color(new A.RgbColorModelHex() { Val = "8064A2" }),
    //                     new A.Accent5Color(new A.RgbColorModelHex() { Val = "4BACC6" }),
    //                     new A.Accent6Color(new A.RgbColorModelHex() { Val = "F79646" }),
    //                     new A.Hyperlink(new A.RgbColorModelHex() { Val = "0000FF" }),
    //                     new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = "800080" })
    //                 ) { Name = "Office" },
    //                 new A.FontScheme(
    //                     new A.MajorFont(
    //                         new A.LatinFont() { Typeface = "Calibri" },
    //                         new A.EastAsianFont() { Typeface = "" },
    //                         new A.ComplexScriptFont() { Typeface = "" } // Fix #2: Added ComplexScriptFont
    //                     ),
    //                     new A.MinorFont(
    //                         new A.LatinFont() { Typeface = "Calibri" },
    //                         new A.EastAsianFont() { Typeface = "" },
    //                         new A.ComplexScriptFont() { Typeface = "" } // Fix #2: Added ComplexScriptFont
    //                     )
    //                 ) { Name = "Office" },
    //                 new A.FormatScheme(
    //                     // Fix #1: FillStyleList needs >= 3 fills
    //                     new A.FillStyleList(
    //                         new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }),
    //                         new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }),
    //                         new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 })
    //                     ),
    //                     // Fix #1: LineStyleList needs >= 3 lines
    //                     new A.LineStyleList(
    //                         new A.Outline(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 })),
    //                         new A.Outline(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 })),
    //                         new A.Outline(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }))
    //                     ),
    //                     // Fix #1: EffectStyleList needs >= 3 effects
    //                     new A.EffectStyleList(
    //                         new A.EffectStyle(new A.EffectList()),
    //                         new A.EffectStyle(new A.EffectList()),
    //                         new A.EffectStyle(new A.EffectList())
    //                     ),
    //                     // Fix #1: BackgroundFillStyleList needs >= 2 fills
    //                     new A.BackgroundFillStyleList(
    //                         new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Background1 }),
    //                         new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Background1 })
    //                     )
    //                 ) { Name = "Office" }
    //             )
    //         );
    //         themePart.Theme.Save();
    //
    //         // 4. ColorMap for Master
    //         ColorMap colorMap = new ColorMap() {
    //             Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1,
    //             Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2,
    //             Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2,
    //             Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4,
    //             Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6,
    //             Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
    //         };
    //
    //         // Fix #3 & #5: Corrected ShapeTree hierarchy placements
    //         ShapeTree masterShapeTree = new ShapeTree(
    //             new NonVisualGroupShapeProperties(
    //                 new NonVisualDrawingProperties() { Id = 1U, Name = "" },
    //                 new NonVisualGroupShapeDrawingProperties(),
    //                 new ApplicationNonVisualDrawingProperties()
    //             ),
    //             new GroupShapeProperties(new A.TransformGroup()) // TransformGroup safely nested
    //         );
    //
    //         slideMasterPart.SlideMaster = new SlideMaster(
    //             new CommonSlideData(masterShapeTree),
    //             colorMap,
    //             new SlideLayoutIdList()
    //         );
    //
    //         // Fix #3, #4 & #5: Slide Layout Structure changes
    //         ShapeTree layoutShapeTree = new ShapeTree(
    //             new NonVisualGroupShapeProperties(
    //                 new NonVisualDrawingProperties() { Id = 1U, Name = "" },
    //                 new NonVisualGroupShapeDrawingProperties(),
    //                 new ApplicationNonVisualDrawingProperties()
    //             ),
    //             new GroupShapeProperties(new A.TransformGroup())
    //         );
    //         slideLayoutPart.SlideLayout = new SlideLayout(
    //             new CommonSlideData(layoutShapeTree)
    //         );
    //         // Fix #4: Add mandatory ColorMapOverride to the layout
    //         slideLayoutPart.SlideLayout.ColorMapOverride = new ColorMapOverride(new A.MasterColorMapping());
    //
    //         // Link Layout to Master Catalog
    //         slideMasterPart.SlideMaster.SlideLayoutIdList.Append(new SlideLayoutId()
    //             { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) });
    //
    //         slideMasterPart.SlideMaster.Save();
    //         slideLayoutPart.SlideLayout.Save();
    //
    //         // 5. Build Root Presentation Mapping
    //         SlideMasterIdList slideMasterIdList = new SlideMasterIdList();
    //         slideMasterIdList.Append(new SlideMasterId() { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) });
    //
    //         SlideIdList slideIdList = new SlideIdList(); // 0 active slides
    //
    //         SlideSize slideSize = new SlideSize() { Cx = (int)emuWidth, Cy = (int)emuHeight, Type = SlideSizeValues.Custom };
    //         NotesSize notesSize = new NotesSize() { Cx = (int)emuHeight, Cy = (int)emuWidth };
    //
    //         presentationPart.Presentation.Append(slideMasterIdList);
    //         presentationPart.Presentation.Append(slideIdList);
    //         presentationPart.Presentation.Append(slideSize);
    //         presentationPart.Presentation.Append(notesSize);
    //
    //         presentationPart.Presentation.Save();
    //     }
    //
    //     stream.Position = 0;
    //     return stream;
    // }

    private static PresentationDocument BuildPresentation(long widthEmu, long heightEmu) {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) {
            Add(zip, "[Content_Types].xml", ContentTypesXml);
            Add(zip, "_rels/.rels", RootRelsXml);
            Add(zip, "ppt/presentation.xml", PresentationXml(widthEmu, heightEmu));
            Add(zip, "ppt/_rels/presentation.xml.rels", PresentationRelsXml);
            Add(zip, "ppt/theme/theme1.xml", ThemeXml);
            Add(zip, "ppt/slideMasters/slideMaster1.xml", SlideMasterXml);
            Add(zip, "ppt/slideMasters/_rels/slideMaster1.xml.rels", SlideMasterRelsXml);
            Add(zip, "ppt/slideLayouts/slideLayout1.xml", SlideLayoutXml);
            Add(zip, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", SlideLayoutRelsXml);
        }

        ms.Position = 0;
        return PresentationDocument.Open(ms, true);
    }

    private static void Add(ZipArchive zip, string name, string xml) {
        using var w = new StreamWriter(zip.CreateEntry(name, CompressionLevel.Fastest).Open(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        w.Write(xml);
    }

    private static string PresentationXml(long cx, long cy) => $"""
                                                                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                                                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                                                                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                                                                  <p:sldMasterIdLst>
                                                                    <p:sldMasterId id="2147483648" r:id="rId1"/>
                                                                  </p:sldMasterIdLst>
                                                                  <p:sldIdLst/>
                                                                  <p:sldSz cx="{cx}" cy="{cy}" type="custom"/>
                                                                  <p:notesSz cx="6858000" cy="9144000"/>
                                                                </p:presentation>
                                                                """;

    private const string ContentTypesXml = """
                                           <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                           <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                                             <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                                             <Default Extension="xml"  ContentType="application/xml"/>
                                             <Override PartName="/ppt/presentation.xml"                ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                                             <Override PartName="/ppt/slideMasters/slideMaster1.xml"   ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                                             <Override PartName="/ppt/slideLayouts/slideLayout1.xml"   ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                                             <Override PartName="/ppt/theme/theme1.xml"                ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                                           </Types>
                                           """;

    private const string RootRelsXml = """
                                       <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                                         <Relationship Id="rId1"
                                           Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                                           Target="ppt/presentation.xml"/>
                                       </Relationships>
                                       """;

    private const string PresentationRelsXml = """
                                               <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                               <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                                                 <Relationship Id="rId1"
                                                   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
                                                   Target="slideMasters/slideMaster1.xml"/>
                                               </Relationships>
                                               """;

    private const string ThemeXml = """
                                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
                                      <a:themeElements>
                                        <a:clrScheme name="Office">
                                          <a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1>
                                          <a:lt1><a:sysClr lastClr="FFFFFF" val="window"/></a:lt1>
                                          <a:dk2><a:srgbClr val="1F3864"/></a:dk2>
                                          <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
                                          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
                                          <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
                                          <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
                                          <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
                                          <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
                                          <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
                                          <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
                                          <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
                                        </a:clrScheme>
                                        <a:fontScheme name="Office">
                                          <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
                                          <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
                                        </a:fontScheme>
                                        <a:fmtScheme name="Office">
                                          <a:fillStyleLst>
                                            <a:noFill/>
                                            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                                            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                                          </a:fillStyleLst>
                                          <a:lnStyleLst>
                                            <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                                            <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                                            <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                                          </a:lnStyleLst>
                                          <a:effectStyleLst>
                                            <a:effectStyle><a:effectLst/></a:effectStyle>
                                            <a:effectStyle><a:effectLst/></a:effectStyle>
                                            <a:effectStyle><a:effectLst/></a:effectStyle>
                                          </a:effectStyleLst>
                                          <a:bgFillStyleLst>
                                            <a:noFill/>
                                            <a:noFill/>
                                            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                                          </a:bgFillStyleLst>
                                        </a:fmtScheme>
                                      </a:themeElements>
                                      <a:objectDefaults/>
                                      <a:extraClrSchemeLst/>
                                    </a:theme>
                                    """;

    private const string SlideMasterXml = """
                                          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                          <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                                            <p:cSld>
                                              <p:spTree>
                                                <p:nvGrpSpPr>
                                                  <p:cNvPr id="1" name=""/>
                                                  <p:cNvGrpSpPr/>
                                                  <p:nvPr/>
                                                </p:nvGrpSpPr>
                                                <p:grpSpPr>
                                                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
                                                          <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
                                                </p:grpSpPr>
                                              </p:spTree>
                                            </p:cSld>
                                            <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
                                                      accent1="accent1" accent2="accent2" accent3="accent3"
                                                      accent4="accent4" accent5="accent5" accent6="accent6"
                                                      hlink="hlink" folHlink="folHlink"/>
                                            <p:sldLayoutIdLst>
                                              <p:sldLayoutId id="2147483649" r:id="rId1"/>
                                            </p:sldLayoutIdLst>
                                            <p:txStyles>
                                              <p:titleStyle/>
                                              <p:bodyStyle/>
                                              <p:otherStyle/>
                                            </p:txStyles>
                                          </p:sldMaster>
                                          """;

    private const string SlideMasterRelsXml = """
                                              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                              <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                                                <Relationship Id="rId1"
                                                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
                                                  Target="../slideLayouts/slideLayout1.xml"/>
                                                <Relationship Id="rId2"
                                                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                                                  Target="../theme/theme1.xml"/>
                                              </Relationships>
                                              """;

    private const string SlideLayoutXml = """
                                          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                          <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                                       type="blank" preserve="1">
                                            <p:cSld name="Blank">
                                              <p:spTree>
                                                <p:nvGrpSpPr>
                                                  <p:cNvPr id="1" name=""/>
                                                  <p:cNvGrpSpPr/>
                                                  <p:nvPr/>
                                                </p:nvGrpSpPr>
                                                <p:grpSpPr>
                                                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
                                                          <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
                                                </p:grpSpPr>
                                              </p:spTree>
                                            </p:cSld>
                                            <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
                                          </p:sldLayout>
                                          """;

    private const string SlideLayoutRelsXml = """
                                              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                              <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                                                <Relationship Id="rId1"
                                                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
                                                  Target="../slideMasters/slideMaster1.xml"/>
                                              </Relationships>
                                              """;
}