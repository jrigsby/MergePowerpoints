using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using PDFiumSharp;

namespace MergePowerpoints;

public static class PdfImageRenderer
{
    private const int TargetDpi = 300;
    private const float SlideSizeInches = 11f;

    public static unsafe MemoryStream RenderPageToSquarePng(byte[] pdfBytes, int pageIndex)
    {
        // Ensure PDFium native binaries are resolvable before first use
        PdfiumLoader.EnsureLoaded();
        using (var doc = new PdfDocument(pdfBytes))
        {
            if (pageIndex < 0 || pageIndex >= doc.Pages.Count)
                throw new ArgumentOutOfRangeException(nameof(pageIndex));

            var page = doc.Pages[pageIndex];
            
            int widthPx = (int)(page.Width * TargetDpi / 72f);
            int heightPx = (int)(page.Height * TargetDpi / 72f);

            using (var pdfBitmap = new PDFiumBitmap(widthPx, heightPx, true))
            {
                page.Render(pdfBitmap);
                
                using (var sourceBmp = new Bitmap(widthPx, heightPx, PixelFormat.Format32bppArgb))
                {
                    var data = sourceBmp.LockBits(new Rectangle(0, 0, widthPx, heightPx), ImageLockMode.WriteOnly, sourceBmp.PixelFormat);
                    
                    IntPtr srcPtr = pdfBitmap.Scan0;
                    IntPtr destPtr = data.Scan0;
                    int byteCount = widthPx * heightPx * 4;
                    
                    Buffer.MemoryCopy(srcPtr.ToPointer(), destPtr.ToPointer(), byteCount, byteCount);
                    
                    sourceBmp.UnlockBits(data);

                    int squareSizePx = (int)(SlideSizeInches * TargetDpi);
                    using (var squareBmp = new Bitmap(squareSizePx, squareSizePx, PixelFormat.Format32bppArgb))
                    {
                        using (var g = Graphics.FromImage(squareBmp))
                        {
                            g.Clear(Color.White);
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                            float scale = Math.Min((float)squareSizePx / widthPx, (float)squareSizePx / heightPx);
                            int finalW = (int)(widthPx * scale);
                            int finalH = (int)(heightPx * scale);

                            int x = (squareSizePx - finalW) / 2;
                            int y = (squareSizePx - finalH) / 2;

                            g.DrawImage(sourceBmp, x, y, finalW, finalH);
                        }

                        var ms = new MemoryStream();
                        squareBmp.Save(ms, ImageFormat.Png);
                        ms.Position = 0;
                        return ms;
                    }
                }
            }
        }
    }

    public static MemoryStream RenderPageToPng(byte[] pdfBytes, int pageIndex, int dpi = 150)
    {
        return RenderPageToSquarePng(pdfBytes, pageIndex);
    }
}
