using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using Tesseract;
using Tesseract.Drawing; // Make sure Tesseract.Drawing NuGet is installed

class Program
{
    // --- WIN32 IMPORTS FOR CAPTURE ---
    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    static void Main(string[] args)
    {
        // 1. SETUP DATA LOOP
        var testCases = new List<Dictionary<string, string>>
        {
            new Dictionary<string, string> { { "Screen", "1040" }, { "ExpectedBottom", "9=Exit" } },
            new Dictionary<string, string> { { "Screen", "SCHA" }, { "ExpectedBottom", "Schedule A" } }
        };

        // 2. CONNECT TO DOS APP (Assuming it's already running for this test)
        Process[] processes = Process.GetProcessesByName("cmd");
        if (processes.Length == 0) return;
        IntPtr hwnd = processes[0].MainWindowHandle;

        // 3. MAIN AUTOMATION LOOP
        foreach (var testCase in testCases)
        {
            string screenName = testCase["Screen"];
            string expectedText = testCase["ExpectedBottom"];

            Console.WriteLine($"\n--- TESTING SCREEN: {screenName} ---");

            // [STUB] Navigation Logic would go here (F3, Type "1040", Enter)
            // InputManager.NavigateTo(screenName); 
            Thread.Sleep(2000); // WAIT for VDI to render the new screen

            // --- OCR CALL STARTS HERE ---
            try
            {
                // A. Capture the current state of the window
                using (Bitmap currentScreen = CaptureWindow(hwnd))
                {
                    // B. Extract and Grade the text from the bottom line
                    string actualBottomText = ExtractBottomLine(currentScreen);
                    
                    Console.WriteLine($"[OCR READ]: '{actualBottomText}'");

                    // C. Verification Logic
                    if (actualBottomText.Contains(expectedText))
                    {
                        Console.WriteLine("✅ PASS: Screen validated.");
                    }
                    else
                    {
                        Console.WriteLine($"❌ FAIL: Expected '{expectedText}', found '{actualBottomText}'");
                        // Optional: Save 'currentScreen' to disk for debugging
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"OCR ERROR: {ex.Message}");
            }
            // --- OCR CALL ENDS HERE ---

            Thread.Sleep(1000); // Pause before next loop iteration
        }
    }

    // --- OCR LOGIC: BOTTOM LINE EXTRACTION ---
    public static string ExtractBottomLine(Bitmap fullCapture)
    {
        // 1. Define Bottom Region (Fixed 30px height from bottom)
        int bottomHeight = 30;
        Rectangle region = new Rectangle(0, fullCapture.Height - bottomHeight, fullCapture.Width, bottomHeight);

        using (Bitmap crop = fullCapture.Clone(region, fullCapture.PixelFormat))
        using (Bitmap processed = PrepareForOcr(crop)) // Apply Grading Filter
        {
            // Debug: Save this specific image to verify the filter works
            // processed.Save(@"C:\Temp\debug_bottom_line.png"); 

            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.LstmOnly))
            {
                // PSM 7 = Single Line (Critical for Status Bars)
                engine.DefaultPageSegMode = PageSegMode.SingleLine;
                engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-.:= ");

                using (var pix = PixConverter.ToPix(processed))
                using (var page = engine.Process(pix))
                {
                    return page.GetText().Trim();
                }
            }
        }
    }

    // --- FILTER LOGIC: GRAYSCALE GRADING ---
    private static Bitmap PrepareForOcr(Bitmap original)
    {
        // 1. Upscale 4x (Standard for DOS pixel fonts)
        Bitmap scaled = new Bitmap(original.Width * 4, original.Height * 4);
        using (Graphics g = Graphics.FromImage(scaled))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.DrawImage(original, 0, 0, scaled.Width, scaled.Height);
        }

        // 2. Apply Grading (Luminance Threshold)
        for (int y = 0; y < scaled.Height; y++)
        {
            for (int x = 0; x < scaled.Width; x++)
            {
                Color c = scaled.GetPixel(x, y);
                // Grading: Calculate brightness (0.0 to 1.0)
                // DOS Text (White/Cyan) is bright (> 0.45). Background (Blue) is dark.
                if (c.GetBrightness() > 0.45f)
                    scaled.SetPixel(x, y, Color.Black); // Make Text Black
                else
                    scaled.SetPixel(x, y, Color.White); // Make Background White
            }
        }
        return scaled;
    }

    // --- HELPER: VDI-SAFE CAPTURE ---
    static Bitmap CaptureWindow(IntPtr hWnd)
    {
        GetWindowRect(hWnd, out RECT rect);
        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;

        if (width <= 0 || height <= 0) return new Bitmap(1, 1);

        Bitmap bmp = new Bitmap(width, height);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            // CopyFromScreen captures exactly what is visible (bypassing VDI black screens)
            g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height));
        }
        return bmp;
    }
}