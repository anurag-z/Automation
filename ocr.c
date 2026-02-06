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

    public static string CaptureAndRead(IntPtr hWnd, string debugDir)
    {
        Directory.CreateDirectory(debugDir);

        // 1. GET COORDINATES
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
            int whiteLineY = -1;
            // Scan middle-column pixels from bottom-up
            for (int y = height - 1; y > height / 2; y--) 
            {
                if (fullContent.GetPixel(width / 2, y).GetBrightness() > 0.85f) 
                {
                    whiteLineY = y;
                    break;
                }
            }

            // 3. CROP BELOW THE BORDER
            int startY = (whiteLineY != -1) ? whiteLineY + 2 : height - 40;
            int captureHeight = height - startY;
            if (captureHeight <= 0) captureHeight = 30; // Fallback

            Rectangle region = new Rectangle(0, startY, width, captureHeight);
            using (Bitmap rawCrop = fullContent.Clone(region, fullContent.PixelFormat))
            {
                rawCrop.Save(Path.Combine(debugDir, "1_Target_Below_Border.png"));

                // 4. CALL THE CALIBRATION PASSES
                return RunCalibrationPasses(rawCrop, debugDir);
            }
        }
    }

    // This is the method that was missing
    private static string RunCalibrationPasses(Bitmap rawCrop, string debugDir)
    {
        float[] gradingLevels = { 0.45f, 0.35f, 0.55f };
        string bestResult = "";

        foreach (float level in gradingLevels)
        {
            using (Bitmap processed = PreProcessImage(rawCrop, level))
            {
                processed.Save(Path.Combine(debugDir, $"2_Processed_Level_{level:0.00}.png"));

                string currentText = RunEngine(processed);

                // Stop if we find indicators like '=', ':' or a long string
                if (currentText.Contains("=") || currentText.Contains(":") || currentText.Length > 10)
                {
                    return currentText;
                }
                if (currentText.Length > bestResult.Length) bestResult = currentText;
            }
        }
        return bestResult;
    }

    private static Bitmap PreProcessImage(Bitmap source, float threshold)
    {
        // Upscale 4x for better letter separation
        Bitmap res = new Bitmap(source.Width * 4, source.Height * 4);
        using (Graphics g = Graphics.FromImage(res))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.DrawImage(source, 0, 0, res.Width, res.Height);
        }

        for (int y = 0; y < res.Height; y++)
        {
            for (int x = 0; x < res.Width; x++)
            {
                Color c = res.GetPixel(x, y);
                // DOS Blue fix: Text (Cyan/White) has high Green component
                bool isText = (c.G > 120) || (c.GetBrightness() > threshold);
                res.SetPixel(x, y, isText ? Color.Black : Color.White);
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