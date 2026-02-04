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
static Bitmap FilterBlueScreen(Bitmap original)
{
    // 1. Scale Up: Use 2x (Not 3x). 
    // Since your DOS font is 24px, 2x makes it 48px tall, which is perfect.
    int scale = 2;
    Bitmap newBmp = new Bitmap(original.Width * scale, original.Height * scale);

    using (Graphics g = Graphics.FromImage(newBmp))
    {
        // CRITICAL: Keep NearestNeighbor to prevent blurry edges
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
        g.DrawImage(original, 0, 0, newBmp.Width, newBmp.Height);
    }

    // 2. Brightness Filter ONLY (No extra bolding)
    for (int y = 0; y < newBmp.Height; y++)
    {
        for (int x = 0; x < newBmp.Width; x++)
        {
            Color c = newBmp.GetPixel(x, y);

            // Calculate Brightness
            // This detects White, Cyan, Yellow, etc.
            int brightness = (int)((c.R * 0.3) + (c.G * 0.59) + (c.B * 0.11));

            // If bright (> 40), make it Black (Text).
            // If dark (< 40), make it White (Background).
            if (brightness > 40) 
            {
                newBmp.SetPixel(x, y, Color.Black); 
            }
            else
            {
                newBmp.SetPixel(x, y, Color.White); 
            }
        }
    }
    
    // Returns clean, separated text without "blobs"
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
