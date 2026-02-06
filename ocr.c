using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Tesseract;
using Tesseract.Drawing; // Required for PixConverter.ToPix()

public static class UltimateOcr
{
    private const string TESS_DATA = @"./tessdata";
    private const string LANGUAGE = "eng";

    /// <summary>
    /// Captures the bottom line of the screen and attempts multiple 
    /// brightness "grading" passes to find the text.
    /// </summary>
    public static string ReadBottomLine(Bitmap fullScreen, string debugDir)
    {
        // 1. ISOLATE: Define the bottom status bar area (approx 35 pixels)
        int cropHeight = 35;
        Rectangle region = new Rectangle(0, fullScreen.Height - cropHeight, fullScreen.Width, cropHeight);

        using (Bitmap rawCrop = fullScreen.Clone(region, fullScreen.PixelFormat))
        {
            // 2. CALIBRATION: Try 3 different "Grading" thresholds to handle VDI color shifts
            float[] gradingLevels = { 0.45f, 0.35f, 0.55f };
            string bestResult = "";

            foreach (float level in gradingLevels)
            {
                using (Bitmap processed = PreProcessImage(rawCrop, level))
                {
                    // Save images for visual debugging
                    string fileName = $"debug_level_{level.ToString("0.00")}.png";
                    processed.Save(Path.Combine(debugDir, fileName));

                    string currentText = RunEngine(processed);

                    // 3. VALIDATION: Check if we found a key character like '=' or a screen digit
                    if (currentText.Contains("=") || currentText.Length > 8)
                    {
                        return currentText; // We found a valid line!
                    }
                    
                    // Fallback to the longest string found if no '=' is present
                    if (currentText.Length > bestResult.Length) bestResult = currentText;
                }
            }
            return bestResult;
        }
    }

    private static Bitmap PreProcessImage(Bitmap source, float threshold)
    {
        // UPSCALE: 3x increase using NearestNeighbor to keep legacy DOS fonts sharp
        int factor = 3;
        Bitmap res = new Bitmap(source.Width * factor, source.Height * factor);
        
        using (Graphics g = Graphics.FromImage(res))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.DrawImage(source, 0, 0, res.Width, res.Height);
        }

        // GRADING: Convert to high-contrast Black & White
        for (int y = 0; y < res.Height; y++)
        {
            for (int x = 0; x < res.Width; x++)
            {
                Color c = res.GetPixel(x, y);
                // Text in DOS is usually Cyan/White (Bright)
                // Background is usually Blue (Dark)
                if (c.GetBrightness() > threshold)
                    res.SetPixel(x, y, Color.Black); // Text -> Black
                else
                    res.SetPixel(x, y, Color.White); // Background -> White
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
                // PSM 7: Treats the image as a single horizontal line (Status bar mode)
                engine.DefaultPageSegMode = PageSegMode.SingleLine;
                
                // Whitelist: Stops Tesseract from guessing random "noise" symbols
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