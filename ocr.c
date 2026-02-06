using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Tesseract;
using Tesseract.Drawing; 

public static class UltimateOcr
{
    // =========================================================
    // 1. WIN32 IMPORTS (For finding and capturing the window)
    // =========================================================
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }

    private const string TESS_DATA = @"./tessdata"; // Path to your tessdata folder
    private const string LANGUAGE = "eng";

    // =========================================================
    // 2. MAIN PUBLIC METHOD: CAPTURE & READ
    // =========================================================
    
    /// <summary>
    /// Captures the window, isolates the bottom line, and reads it using multi-pass grading.
    /// </summary>
    public static string CaptureAndRead(IntPtr hWnd, string debugDir)
    {
        // STEP 1: CAPTURE
        using (Bitmap fullScreen = CaptureWindow(hWnd))
        {
            if (fullScreen == null) return "ERROR_CAPTURE_FAILED";

            // STEP 2: ISOLATE BOTTOM LINE (Approx 35 pixels)
            int cropHeight = 35;
            Rectangle region = new Rectangle(0, fullScreen.Height - cropHeight, fullScreen.Width, cropHeight);

            using (Bitmap rawCrop = fullScreen.Clone(region, fullScreen.PixelFormat))
            {
                // STEP 3: CALIBRATION LOOP (Try 3 grading levels)
                float[] gradingLevels = { 0.45f, 0.35f, 0.55f };
                string bestResult = "";

                foreach (float level in gradingLevels)
                {
                    using (Bitmap processed = PreProcessImage(rawCrop, level))
                    {
                        // Save debug image
                        string fileName = $"debug_level_{level.ToString("0.00")}.png";
                        processed.Save(Path.Combine(debugDir, fileName));

                        // STEP 4: RUN TESSERACT
                        string currentText = RunEngine(processed);

                        // Validation: If we found a known marker, stop immediately.
                        if (currentText.Contains("=") || currentText.Length > 8)
                        {
                            return currentText; 
                        }
                        
                        // Keep the longest result just in case
                        if (currentText.Length > bestResult.Length) bestResult = currentText;
                    }
                }
                return bestResult;
            }
        }
    }

    // =========================================================
    // 3. PRIVATE HELPER METHODS
    // =========================================================

    private static Bitmap CaptureWindow(IntPtr hWnd)
    {
        GetWindowRect(hWnd, out RECT rect);
        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;

        if (width <= 0 || height <= 0) return null;

        Bitmap bmp = new Bitmap(width, height);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            // CopyFromScreen is best for VDI as it captures "what you see"
            g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height));
        }
        return bmp;
    }

    private static Bitmap PreProcessImage(Bitmap source, float threshold)
    {
        // UPSCALE 3x (NearestNeighbor for sharp pixels)
        int factor = 3;
        Bitmap res = new Bitmap(source.Width * factor, source.Height * factor);
        
        using (Graphics g = Graphics.FromImage(res))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.DrawImage(source, 0, 0, res.Width, res.Height);
        }

        // GRADE (Luminance Threshold)
        for (int y = 0; y < res.Height; y++)
        {
            for (int x = 0; x < res.Width; x++)
            {
                Color c = res.GetPixel(x, y);
                // DOS Text is Bright (White/Cyan) vs Dark Blue BG
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
                engine.DefaultPageSegMode = PageSegMode.SingleLine;
                // Strict whitelist to reduce "CO" noise
                engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-.:= ");

                using (var pix = PixConverter.ToPix(img))
                using (var page = engine.Process(pix))
                {
                    return page.GetText().Trim();
                }
            }
        }
        catch (Exception ex)
        {
            return $"ERR: {ex.Message}";
        }
    }
}