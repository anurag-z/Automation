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
            for (int y = height - 1; y > height / 2; y--) 
            {
                // Scan the middle column for the bright white border
                if (fullContent.GetPixel(width / 2, y).GetBrightness() > 0.85f) 
                {
                    whiteLineY = y;
                    break;
                }
            }

            // 3. CROP BELOW THE BORDER
            // Start 3 pixels below the white line to ensure no "line noise" enters the OCR
            int startY = (whiteLineY != -1) ? whiteLineY + 3 : height - 35;
            int captureHeight = height - startY;
            if (captureHeight <= 0) captureHeight = 30;

            Rectangle region = new Rectangle(0, startY, width, captureHeight);
            using (Bitmap rawCrop = fullContent.Clone(region, fullContent.PixelFormat))
            {
                rawCrop.Save(Path.Combine(debugDir, "1_Target_Crop.png"));
                return RunCalibrationPasses(rawCrop, debugDir);
            }
        }
    }

    private static string RunCalibrationPasses(Bitmap rawCrop, string debugDir)
    {
        // Try different threshold levels. 0.40 is usually the "sweet spot" for Cyan-on-Blue.
        float[] gradingLevels = { 0.40f, 0.30f, 0.50f };
        string bestResult = "";

        foreach (float level in gradingLevels)
        {
            using (Bitmap processed = PreProcessImage(rawCrop, level))
            {
                processed.Save(Path.Combine(debugDir, $"2_Final_For_Tesseract_{level:0.00}.png"));

                string currentText = RunEngine(processed);

                // If the result contains our key navigation markers, we found it!
                if (currentText.Contains("=") || currentText.Contains("-")) return currentText;
                if (currentText.Length > bestResult.Length) bestResult = currentText;
            }
        }
        return bestResult;
    }

    private static Bitmap PreProcessImage(Bitmap source, float threshold)
{
    // 1. UPSCALE 4x
    // Larger images allow the erosion logic to be more precise.
    Bitmap res = new Bitmap(source.Width * 4, source.Height * 4);
    using (Graphics g = Graphics.FromImage(res))
    {
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
        g.DrawImage(source, 0, 0, res.Width, res.Height);
    }

    // 2. CREATE A BITMAP MATRIX (Based on your 0.50f threshold)
    bool[,] matrix = new bool[res.Width, res.Height];
    for (int y = 0; y < res.Height; y++)
    {
        for (int x = 0; x < res.Width; x++)
        {
            Color c = res.GetPixel(x, y);
            // Current working logic: threshold = 0.50f
            matrix[x, y] = (c.G > 115) || (c.GetBrightness() > threshold);
        }
    }

    // 3. APPLY EROSION (The "Zero Fix")
    // We check the 4 immediate neighbors (Up, Down, Left, Right).
    for (int y = 1; y < res.Height - 1; y++)
    {
        for (int x = 1; x < res.Width - 1; x++)
        {
            if (matrix[x, y]) // If it's a "Black" (text) pixel
            {
                int neighbors = 0;
                if (matrix[x - 1, y]) neighbors++;
                if (matrix[x + 1, y]) neighbors++;
                if (matrix[x, y - 1]) neighbors++;
                if (matrix[x, y + 1]) neighbors++;

                // A pixel in a thick line (like the wall of a 0) has 3-4 neighbors.
                // A pixel in a thin diagonal dash usually only has 2.
                // By removing pixels with < 3 neighbors, we break the dash.
                if (neighbors < 3) 
                    res.SetPixel(x, y, Color.White); // Erase noise/dash
                else
                    res.SetPixel(x, y, Color.Black); // Keep solid text
            }
            else
            {
                res.SetPixel(x, y, Color.White);
            }
        }
    }
    return res;
}

    private static string RunEngine(Bitmap img)
{
    try
    {
        // 1. Convert Bitmap to a format Tesseract 5.0 loves (32bpp Argb)
        // and save to memory to "normalize" the DPI.
        using (MemoryStream ms = new MemoryStream())
        {
            img.Save(ms, ImageFormat.Png);
            ms.Position = 0;

            using (var engine = new TesseractEngine(TESS_DATA, LANGUAGE, EngineMode.LstmOnly))
            {
                // 2. These settings mirror the standalone Tesseract command line
                engine.SetVariable("user_defined_dpi", "300");
                engine.DefaultPageSegMode = PageSegMode.SingleLine;
                engine.SetVariable("tessedit_char_whitelist", "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-.:= ");

                // 3. Load from the "File-Like" stream
                using (var pix = Pix.LoadFromMemory(ms.ToArray()))
                {
                    using (var page = engine.Process(pix))
                    {
                        string result = page.GetText().Trim();
                        return result;
                    }
                }
            }
        }
    }
    catch (Exception ex)
    {
        return $"ERROR: {ex.Message}";
    }
}

}
