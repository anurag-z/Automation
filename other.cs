using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using Tesseract;
using Tesseract.Drawing;

class DosAutomation
{
    // ---------------- WIN32 ----------------
    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    const int SW_RESTORE = 9;

    static TesseractEngine engine;

    static void Main()
    {
        string tessDataPath = @"./tessdata";
        string debugFolder = @"C:\Temp\OCR_Debug";
        Directory.CreateDirectory(debugFolder);

        // --- INIT OCR ONCE ---
        engine = new TesseractEngine(tessDataPath, "eng", EngineMode.LstmOnly);
        engine.DefaultPageSegMode = PageSegMode.SingleLine;
        engine.SetVariable("tessedit_char_whitelist",
            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-.:=,() ");

        try
        {
            // --- LAUNCH DOS ---
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                WorkingDirectory = @"C:\10405",
                Arguments = "/k fads",
                UseShellExecute = true
            };

            Process p = Process.Start(psi);

            // --- WAIT FOR WINDOW ---
            IntPtr hwnd = IntPtr.Zero;
            for (int i = 0; i < 15; i++)
            {
                p.Refresh();
                hwnd = p.MainWindowHandle;
                if (hwnd != IntPtr.Zero) break;
                Thread.Sleep(1000);
            }

            if (hwnd == IntPtr.Zero)
                throw new Exception("DOS window not found.");

            ShowWindow(hwnd, SW_RESTORE);
            SetForegroundWindow(hwnd);
            Thread.Sleep(1000);

            using (Bitmap screen = CaptureWindow(hwnd))
            {
                var result = ReadTargetedAreas(screen, debugFolder);

                Console.WriteLine("\n--- OCR RESULT ---");
                Console.WriteLine("HIGHLIGHTED ROW:");
                Console.WriteLine(result.HighlightedRow);
                Console.WriteLine("\nBOTTOM LINE:");
                Console.WriteLine(result.BottomLine);
                Console.WriteLine("------------------");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR: " + ex.Message);
        }
        finally
        {
            engine?.Dispose();
        }
    }

    // ---------------- OCR LOGIC ----------------

    static (string HighlightedRow, string BottomLine) ReadTargetedAreas(Bitmap bmp, string debug)
    {
        int highlightY = DetectHighlightRow(bmp);
        int lineHeight = bmp.Height / 25; // DOS = 25 rows

        string highlight = "";
        string bottom = "";

        if (highlightY > 0)
        {
            Rectangle highlightRect = new Rectangle(
                0,
                Math.Max(0, highlightY - lineHeight / 2),
                bmp.Width,
                lineHeight
            );

            highlight = ExtractText(bmp, highlightRect,
                Path.Combine(debug, "highlight.png"));
        }

        Rectangle bottomRect = new Rectangle(
            0,
            bmp.Height - lineHeight,
            bmp.Width,
            lineHeight
        );

        bottom = ExtractText(bmp, bottomRect,
            Path.Combine(debug, "bottom.png"));

        return (highlight, bottom);
    }

    static int DetectHighlightRow(Bitmap bmp)
    {
        int bestRow = -1;
        double bestScore = 0;

        for (int y = 0; y < bmp.Height - 40; y++)
        {
            double score = 0;

            for (int x = 0; x < bmp.Width; x += 6)
            {
                Color c = bmp.GetPixel(x, y);
                score += (0.2126 * c.R + 0.7152 * c.G + 0.0722 * c.B);
            }

            if (score > bestScore)
            {
                bestScore = score;
                bestRow = y;
            }
        }
        return bestRow;
    }

    static string ExtractText(Bitmap source, Rectangle region, string debugPath)
    {
        using (Bitmap crop = source.Clone(region, source.PixelFormat))
        using (Bitmap processed = CleanImageForOcr(crop))
        {
            processed.Save(debugPath, ImageFormat.Png);

            using (var pix = PixConverter.ToPix(processed))
            using (var page = engine.Process(pix))
            {
                return page.GetText().Trim();
            }
        }
    }

    static Bitmap CleanImageForOcr(Bitmap input)
    {
        Bitmap scaled = new Bitmap(input.Width * 4, input.Height * 4);

        using (Graphics g = Graphics.FromImage(scaled))
        {
            g.InterpolationMode =
                System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.DrawImage(input, 0, 0, scaled.Width, scaled.Height);
        }

        for (int y = 0; y < scaled.Height; y++)
        {
            for (int x = 0; x < scaled.Width; x++)
            {
                Color c = scaled.GetPixel(x, y);
                scaled.SetPixel(
                    x, y,
                    c.GetBrightness() < 0.5f ? Color.Black : Color.White
                );
            }
        }
        return scaled;
    }

    // ---------------- SCREEN CAPTURE ----------------

    static Bitmap CaptureWindow(IntPtr hWnd)
    {
        GetWindowRect(hWnd, out RECT rect);

        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;

        Bitmap bmp = new Bitmap(width, height);

        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(rect.Left, rect.Top, 0, 0,
                new Size(width, height));
        }
        return bmp;
    }
}
