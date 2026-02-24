using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace MergePowerpoints;

public static class PptNormalizer {
    private const float SlideSizePt = 11f * 72f; // 792 pt (11 in)

    // Normalize a PPTX stream: set slide size to 11x11 and scale/center content per slide.
    // Returns a new stream positioned at 0.
    public static MemoryStream NormalizePptxStreamToSquare(Stream inputPptx) {
        if (inputPptx == null) throw new ArgumentNullException(nameof(inputPptx));
        if (inputPptx.CanSeek) inputPptx.Position = 0;

        using var pres = new Presentation();
        pres.LoadFromStream(inputPptx, FileFormat.Pptx2013);

        // Set deck slide size to 11x11
        try {
            pres.SlideSize.Type = SlideSizeType.Custom;
            pres.SlideSize.Size = new SizeF(SlideSizePt, SlideSizePt);
        }
        catch (Exception ex) {
            // If setting size at the deck level fails due to internal Spire.Presentation issues
            // (e.g., trying to auto-scale corrupt internal objects), we'll try to proceed.
            // The NormalizeSlide method will still attempt to position shapes within the 11x11 boundary.
            Console.WriteLine($"[WARNING] Failed to set global slide size: {ex.Message}");
        }

        foreach (ISlide slide in pres.Slides) {
            NormalizeSlide(slide);
        }

        var output = new MemoryStream();
        pres.SaveToFile(output, FileFormat.Pptx2013);
        output.Position = 0;
        return output;
    }

    private static void NormalizeSlide(ISlide slide) {
        if (slide == null) return;

        // Ensure Shapes collection is accessible
        var shapes = slide.Shapes;
        if (shapes == null || shapes.Count == 0) return;

        // Compute content bounding box
        float minX = float.MaxValue, minY = float.MaxValue;
        float maxX = float.MinValue, maxY = float.MinValue;

        for (int i = 0; i < shapes.Count; i++) {
            IShape s;
            try {
                s = shapes[i];
            }
            catch {
                continue;
            }

            // Some shapes may not be visible or may be placeholders; include all visible shapes
            if (s == null) continue;

            try {
                float l = s.Left;
                float t = s.Top;
                float w = s.Width;
                float h = s.Height;
                minX = Math.Min(minX, l);
                minY = Math.Min(minY, t);
                maxX = Math.Max(maxX, l + w);
                maxY = Math.Max(maxY, t + h);
            }
            catch {
                // Skip shapes with invalid properties
                continue;
            }
        }

        if (minX == float.MaxValue) return; // nothing meaningful

        float bboxW = Math.Max(1f, maxX - minX);
        float bboxH = Math.Max(1f, maxY - minY);

        // Scale to fit square slide while preserving aspect ratio
        float scale = Math.Min(SlideSizePt / bboxW, SlideSizePt / bboxH);
        float targetW = bboxW * scale;
        float targetH = bboxH * scale;

        // Compute offsets to center
        float offsetX = (SlideSizePt - targetW) / 2f;
        float offsetY = (SlideSizePt - targetH) / 2f;

        // Apply transform to all shapes
        for (int i = 0; i < shapes.Count; i++) {
            IShape s;
            try {
                s = shapes[i];
            }
            catch {
                continue;
            }

            if (s == null) continue;

            try {
                float newLeft = (s.Left - minX) * scale + offsetX;
                float newTop = (s.Top - minY) * scale + offsetY;
                float newWidth = s.Width * scale;
                float newHeight = s.Height * scale;

                s.Left = newLeft;
                s.Top = newTop;
                s.Width = newWidth;
                s.Height = newHeight;

                // Remove borders for auto-shapes to avoid thin outline after scaling
                if (s is IAutoShape auto) {
                    if (auto.Line != null) {
                        auto.Line.FillType = FillFormatType.None;
                    }
                }
            }
            catch {
                // Skip shapes that can't be transformed
                continue;
            }
        }

        // Optional: set background to white for consistency (guard against null background/fill on some imports)
        try {
            if (slide.SlideBackground != null) {
                slide.SlideBackground.Type = BackgroundType.Custom;
                if (slide.SlideBackground.Fill != null) {
                    slide.SlideBackground.Fill.FillType = FillFormatType.Solid;
                    slide.SlideBackground.Fill.SolidColor.Color = Color.White;
                }
            }
        }
        catch {
            // Ignore background errors; normalization of shapes is the key requirement.
        }
    }
}