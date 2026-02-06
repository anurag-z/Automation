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
        Directory.CreateDirectory(debugDir);

        // 1. GET ACCURATE COORDINATES (Content area only)
        GetClientRect(hWnd, out RECT clientRect);
        POINT topLeft = new POINT { X = 0, Y = 0 };
        ClientToScreen(hWnd, ref topLeft);

        int width = clientRect.Right - clientRect.Left;
        int height = clientRect.Bottom - clientRect.Top;

        if (width <= 0 || height <= 0) return "ERROR_INVALID_WINDOW_SIZE";

        using (Bitmap fullContent = new Bitmap(width, height))
        {
            // 2. CAPTURE THE SCREEN
            using (Graphics g = Graphics.FromImage(fullContent))
            {
                g.CopyFromScreen(topLeft.X, topLeft.Y, 0, 0, new Size(width, height));
            }
            fullContent.Save(Path.Combine(debugDir, "0_Content_Area_Only.png"));

            // 3. ISOLATE THE BOTTOM STRIP 
            // We take a slightly larger strip (60px) and offset it from the absolute bottom
            int stripHeight = 60; 
            int bottomGap = 5; 
            Rectangle cropRegion = new Rectangle(0, height - stripHeight - bottomGap, width, stripHeight);

            using (Bitmap rawCrop = fullContent.Clone(cropRegion, fullContent.PixelFormat))
            {
                rawCrop.Save(Path.Combine(debugDir, "1_Bottom_Crop_Raw.png"));

                // 4. MULTI-PASS OCR GRADING
                float[] gradingLevels = { 0.35f, 0.45f, 0.55f };
                string bestResult = "";

                foreach (float level in gradingLevels)
                {
                    using (Bitmap processed = PreProcessImage(rawCrop, level))
                    {
                        string fileName = $"2_Processed_Level_{level.ToString("0.00")}.png";
                        processed.Save(Path.Combine(debugDir, fileName));

                        string currentText = RunEngine(processed);

                        // If we see typical DOS status indicators, return immediately
                        if (currentText.Contains("=") || currentText.Contains(":") || currentText.Length > 10)
                        {
                            return currentText;
                        }

                        if (currentText.Length > bestResult.Length) bestResult = currentText;
                    }
                }
                return bestResult;
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