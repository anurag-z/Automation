using System;
using System.Diagnostics;
using System.Drawing; // Requires System.Drawing.Common
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using Tesseract; // Requires Tesseract NuGet

class Program
{
    // --- SETUP ---
    [DllImport("user32.dll")]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    const int KEYEVENTF_KEYUP = 0x0002;
    const int KEYEVENTF_SCANCODE = 0x0008;
    const byte SC_F3 = 0x3D;

    static void Main()
    {
        ProcessStartInfo processInfo = new ProcessStartInfo();
        processInfo.FileName = "cmd.exe";
        processInfo.WorkingDirectory = @"c:\10405";
        processInfo.Arguments = @"/k fads";
        processInfo.UseShellExecute = true;

        Process p = Process.Start(processInfo);
        Thread.Sleep(2000); 

        try
        {
            Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
            Thread.Sleep(500);
            
            Console.WriteLine("Sending F3...");
            PressKey(SC_F3);
            Thread.Sleep(2000);

            // 1. CAPTURE
            Console.WriteLine("Taking screenshot...");
            using (Bitmap original = CaptureWindow(p.MainWindowHandle))
            {
                // 2. PROCESS (Fix for Blue Screen)
                Console.WriteLine("Applying Blue-Screen Filter...");
                using (Bitmap processed = FilterBlueScreen(original))
                {
                    // Save debug image to check logic
                    string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    processed.Save(Path.Combine(desktop, "debug_processed_blue.png"), ImageFormat.Png);

                    // 3. CONVERT (Bitmap -> File -> Pix)
                    // We save to a temp file because Tesseract cannot read Bitmap directly
                    string tempFile = "temp_ocr.tif";
                    processed.Save(tempFile, ImageFormat.Tiff);

                    try
                    {
                        using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                        {
                            // Load the temp file as 'Pix'
                            using (var img = Pix.LoadFromFile(tempFile)) 
                            {
                                using (var page = engine.Process(img, PageSegMode.SingleBlock))
                                {
                                    string text = page.GetText();
                                    Console.WriteLine("--- FOUND TEXT ---");
                                    Console.WriteLine(text);
                                    Console.WriteLine("------------------");

                                    // Verify
                                    string cleanText = text.ToUpper().Replace(" ", "");
                                    
                                    if (cleanText.Contains("F7"))
                                        Console.WriteLine("[PASS] 'F7' found.");
                                    else if (cleanText.Contains("FEDERAL"))
                                        Console.WriteLine("[PASS] Header found.");
                                    else
                                        Console.WriteLine("[FAIL] Target text not found.");
                                }
                            }
                        }
                    }
                    finally
                    {
                        // Cleanup temp file
                        if (File.Exists(tempFile)) File.Delete(tempFile);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // --- HELPER: Blue Screen Filter ---
   // --- IMPROVED FILTER: COLOR DISTANCE ---
   // --- FINAL FILTER: CHANNEL CHECK ---
   // --- FINAL ROBUST FILTER: THE "RED+GREEN" SUM ---
  // --- REPLACE YOUR EXISTING FILTER METHOD WITH THIS ---
// REPLACE YOUR FILTER METHOD WITH THIS ONE
static Bitmap FilterBlueScreen(Bitmap original)
{
    // 1. Scale Up (2x) 
    // We use 2x because your Font Size is 24.
    // This creates a nice large image (approx 1600px wide) for OCR.
    int scale = 2;
    int padding = 20; // Adds white border so text doesn't touch the edge
    int w = original.Width * scale;
    int h = original.Height * scale;

    Bitmap newBmp = new Bitmap(w + (padding * 2), h + (padding * 2));

    using (Graphics g = Graphics.FromImage(newBmp))
    {
        g.Clear(Color.White); // Start with a clean white sheet

        // MANDATORY: NearestNeighbor keeps pixels square and sharp
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
        
        // Draw the original image inside the padding
        g.DrawImage(original, padding, padding, w, h);
    }

    // 2. The "Green or Red" Filter
    for (int y = 0; y < newBmp.Height; y++)
    {
        for (int x = 0; x < newBmp.Width; x++)
        {
            Color c = newBmp.GetPixel(x, y);

            // LOGIC EXPLAINED:
            // - Blue Background: (R:0, G:0, B:170) -> Has NO Red or Green.
            // - Cyan Text:       (R:0, G:255, B:255) -> Has GREEN.
            // - White Text:      (R:255, G:255, B:255) -> Has RED and GREEN.
            
            // If the pixel has ANY significant Red OR Green, it must be text.
            // We ignore Blue completely because everything has Blue in it.
            
            if (c.R > 60 || c.G > 60) 
            {
                newBmp.SetPixel(x, y, Color.Black); // Keep it (Ink)
            }
            else
            {
                newBmp.SetPixel(x, y, Color.White); // Remove it (Paper)
            }
        }
    }
    
    return newBmp;
}
    static Bitmap CaptureWindow(IntPtr handle)
    {
        RECT rect;
        GetWindowRect(handle, out rect);
        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;

        Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(rect.Left, rect.Top, 0, 0, bmp.Size, CopyPixelOperation.SourceCopy);
        }
        return bmp;
    }

    static void PressKey(byte scanCode)
    {
        keybd_event(0, scanCode, KEYEVENTF_SCANCODE, UIntPtr.Zero);
        Thread.Sleep(100);
        keybd_event(0, scanCode, KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
