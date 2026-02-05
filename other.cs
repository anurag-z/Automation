using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using Tesseract;

class DosAutomation
{
    // --- WIN32 API for Window Management ---
    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    const int SW_RESTORE = 9;

    static void Main()
    {
        // 1. Setup paths
        string tessDataPath = @"./tessdata"; // Ensure this folder contains eng.traineddata
        string debugFolder = @"C:\Temp\OCR_Debug";
        Directory.CreateDirectory(debugFolder);

        try
        {
            // 2. Launch or Find the DOS App
            Process[] processes = Process.GetProcessesByName("cmd");
            if (processes.Length == 0) { Console.WriteLine("DOS App (cmd) not running."); return; }
            IntPtr hwnd = processes[0].MainWindowHandle;

            // Bring to front - Essential for VDI pixel capture
            ShowWindow(hwnd, SW_RESTORE);
            SetForegroundWindow(hwnd);
            Thread.Sleep(1000); // Wait for VDI redraw

            // 3. Capture and Process
            using (Bitmap fullScreen = CaptureWindow(hwnd))
            {
                var result = ReadTargetedAreas(fullScreen, tessDataPath, debugFolder);

                Console.WriteLine("\n--- OCR RESULTS ---");
                Console.WriteLine($"BOTTOM LINE: {result.BottomLine}");
                Console.WriteLine($"HIGHLIGHTED: {result.HighlightedRow}");
                Console.WriteLine("-------------------\n");

                // 4. Verification Example
                if (result.BottomLine.Contains("9=Exit"))
                {
                    Console.WriteLine("Verification Success: On Main Menu.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // --- TARGETED READING LOGIC ---
    static (string HighlightedRow, string BottomLine) ReadTargetedAreas(Bitmap bmp, string dataPath, string debug)
    {
        string highlightedText = "";
        string bottomText = "";

        // Area A: Find Highlighted Row (Dynamic Y-Axis)
        int midX = bmp.Width / 2;
        int startY = -1, endY = -1;
        for (int y = 0; y < bmp.Height - 40; y++)
        {
            Color c = bmp.GetPixel(midX, y);
            // Detect non-standard blue (Highlight color)
            if (c.B > 150 && c.R < 100 && c.G < 180) 
            {
                if (startY == -1) startY = y;
                endY = y;
            }
        }

        // OCR Highlighted Line
        if (startY != -1)
        {
            Rectangle rect = new Rectangle(0, startY, bmp.Width, (endY - startY) + 2);
            highlightedText = ExtractText(bmp, rect, dataPath, Path.Combine(debug, "highlight.png"));
        }

        // Area B: Bottom Line (Fixed Y-Axis)
        Rectangle bottomRect = new Rectangle(0, bmp.Height - 30, bmp.Width, 30);
        bottomText = ExtractText(bmp, bottomRect, dataPath, Path.Combine(debug, "bottom.png"));

        return (highlightedText, bottomText);
    }

    static string ExtractText(Bitmap source, Rectangle region, string dataPath, string debugPath)
    {
        using (Bitmap crop = source.Clone(region, source.PixelFormat))
        using (Bitmap processed = CleanImageForOcr(crop))
        {
            processed.Save(debugPath); // Save for visual verification

            using (var engine = new TesseractEngine(dataPath, "eng", EngineMode.LstmOnly))
            {
                // PSM 7: Treat as a single text line for maximum accuracy
                engine.DefaultPageSegMode = PageSegMode.SingleLine;
                engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-.:= ");

                using (var page = engine.Process(processed))
                {
                    return page.GetText().Trim();
                }
            }
        }
    }

    static Bitmap CleanImageForOcr(Bitmap part)
    {
        // Scale 4x using NearestNeighbor to keep DOS pixels sharp
        Bitmap scaled = new Bitmap(part.Width * 4, part.Height * 4);
        using (Graphics g = Graphics.FromImage(scaled))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.DrawImage(part, 0, 0, scaled.Width, scaled.Height);
        }

        // Binarization (Luminance Threshold)
        for (int y = 0; y < scaled.Height; y++)
        {
            for (int x = 0; x < scaled.Width; x++)
            {
                Color c = scaled.GetPixel(x, y);
                scaled.SetPixel(x, y, c.GetBrightness() > 0.45f ? Color.Black : Color.White);
            }
        }
        return scaled;
    }

    static Bitmap CaptureWindow(IntPtr hWnd)
    {
        GetWindowRect(hWnd, out RECT rect);
        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;

        Bitmap bmp = new Bitmap(width, height);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            // Direct screen copy bypasses VDI hardware acceleration black-screens
            g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height));
        }
        return bmp;
    }
}