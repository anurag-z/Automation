using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Tesseract;
using Tesseract.Drawing; 

public static class UltimateOcr
{
    // --- WIN32 API FOR COORDINATES ---
    [DllImport("user32.dll")]
    static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X; public int Y; }

    private const string TESS_DATA = @"./tessdata";
    private const string LANGUAGE = "eng";

    /// <summary>
    /// Captures the DOS window, isolates the bottom line, and performs OCR.
    /// </summary>
    public static string CaptureAndRead(IntPtr hWnd, string debugDir)
{
    // 1. GET THE FULL CONTENT AREA
    GetClientRect(hWnd, out RECT clientRect);
    POINT topLeft = new POINT { X = 0, Y = 0 };
    ClientToScreen(hWnd, ref topLeft);

    int width = clientRect.Right - clientRect.Left;
    int height = clientRect.Bottom - clientRect.Top;

    using (Bitmap fullContent = new Bitmap(width, height))
    {
        using (Graphics g = Graphics.FromImage(fullContent))
        {
            g.CopyFromScreen(topLeft.X, topLeft.Y, 0, 0, new Size(width, height));
        }

        // 2. FIND THE WHITE BORDER LINE
        // We scan from the bottom upwards to find the first solid line of light pixels
        int whiteLineY = -1;
        for (int y = height - 1; y > height / 2; y--) // Scan bottom half only
        {
            Color pixel = fullContent.GetPixel(width / 2, y); // Check the middle of the row
            // If the pixel is very bright (White/Cyan border), we found our line
            if (pixel.GetBrightness() > 0.8f) 
            {
                whiteLineY = y;
                break;
            }
        }

        // 3. DEFINE THE CROP BASED ON THE BORDER
        // If we found the line, we start 2 pixels BELOW it. 
        // If not found, we fallback to the bottom 40 pixels.
        int startY = (whiteLineY != -1) ? whiteLineY + 2 : height - 40;
        int captureHeight = (whiteLineY != -1) ? (height - startY) : 35;

        // Safety check to ensure we don't crop outside the image
        if (startY + captureHeight > height) captureHeight = height - startY;

        Rectangle region = new Rectangle(0, startY, width, captureHeight);

        using (Bitmap rawCrop = fullContent.Clone(region, fullContent.PixelFormat))
        {
            rawCrop.Save(Path.Combine(debugDir, "1_Target_Below_Border.png"));

            // 4. MULTI-PASS OCR (Same as before)
            return RunCalibrationPasses(rawCrop, debugDir);
        }
    }
}

    private static Bitmap PreProcessImage(Bitmap source, float threshold)
    {
        // UPSCALE 3x (Crucial for Tesseract to see DOS pixel fonts correctly)
        Bitmap res = new Bitmap(source.Width * 3, source.Height * 3);
        using (Graphics g = Graphics.FromImage(res))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.DrawImage(source, 0, 0, res.Width, res.Height);
        }

        // BINARIZATION (Grading)
        for (int y = 0; y < res.Height; y++)
        {
            for (int x = 0; x < res.Width; x++)
            {
                Color c = res.GetPixel(x, y);
                // DOS text is usually brighter than the background
                if (c.GetBrightness() > threshold)
                    res.SetPixel(x, y, Color.Black); // Text
                else
                    res.SetPixel(x, y, Color.White); // Background
            }
        }
        return res;
    }

    private static string RunEngine(Bitmap img)
    {
        try
        {
            using (var engine = new TesseractEngine(TESS_DATA, LANGUAGE, EngineMode.LstmOnly))
            {
                // Single line mode is most accurate for status bars
                engine.DefaultPageSegMode = PageSegMode.SingleLine;
                engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-.:= ");

                using (var pix = PixConverter.ToPix(img))
                using (var page = engine.Process(pix))
                {
                    return page.GetText().Trim();
                }
            }
        }
        catch { return ""; }
    }
}