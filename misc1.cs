using System;
using System.Diagnostics;
using System.Drawing; // System.Drawing.Common
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using Tesseract; // Tesseract NuGet

class Program
{
    // --- IMPORTS ---
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
        try
        {
            // 1. LAUNCH
            ProcessStartInfo processInfo = new ProcessStartInfo();
            processInfo.FileName = "cmd.exe";
            processInfo.WorkingDirectory = @"c:\10405";
            processInfo.Arguments = @"/k fads"; 
            processInfo.UseShellExecute = true;

            Process p = Process.Start(processInfo);
            Thread.Sleep(2000); 

            Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
            Thread.Sleep(500);
            
            Console.WriteLine("Sending F3...");
            PressKey(SC_F3);
            Thread.Sleep(2000);

            Console.WriteLine("Taking screenshot...");
            using (Bitmap original = CaptureWindow(p.MainWindowHandle))
            {
                // 2. PROCESS (Using the "Precision" Method)
                Console.WriteLine("Processing (Precision Mode)...");
                using (Bitmap processed = FilterPrecision(original))
                {
                    // DEBUG: Save to desktop to VERIFY image is White with Sharp Black text
                    string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string debugPath = Path.Combine(desktop, "debug_precision.png");
                    processed.Save(debugPath, ImageFormat.Png);
                    Console.WriteLine($"Image Saved: {debugPath} <--- OPEN THIS");

                    // 3. OCR
                    string tempFile = "temp_ocr.tif";
                    processed.Save(tempFile, ImageFormat.Tiff);

                    try
                    {
                        using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                        {
                            // Whitelist to force clean reading
                            engine.SetVariable("tessedit_char_whitelist", "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ,.-() ");
                            
                            using (var img = Pix.LoadFromFile(tempFile)) 
                            {
                                // SingleBlock is best for this menu layout
                                using (var page = engine.Process(img, PageSegMode.SingleBlock))
                                {
                                    string text = page.GetText();
                                    Console.WriteLine("\n--- EXTRACTED TEXT ---");
                                    Console.WriteLine(text);
                                    Console.WriteLine("----------------------");
                                }
                            }
                        }
                    }
                    finally
                    {
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

    static Bitmap FilterBlueScreen(Bitmap original)
{
    // STEP 1: Pass One - Safe Capture
    // We create a base layer using a standard threshold.
    Bitmap baseLayer = new Bitmap(original.Width, original.Height);
    using (Graphics g = Graphics.FromImage(baseLayer)) 
    { 
        g.Clear(Color.White); // FORCE WHITE BACKGROUND
    }

    for (int y = 0; y < original.Height; y++)
    {
        for (int x = 0; x < original.Width; x++)
        {
            Color c = original.GetPixel(x, y);
            // Brightness check
            int b = (int)((c.R * 0.3) + (c.G * 0.59) + (c.B * 0.11));

            // Threshold 45 is the "Sweet Spot" for your DOS screen.
            if (b > 45) baseLayer.SetPixel(x, y, Color.Black);
        }
    }

    // STEP 2: Pass Two - Precision Connection
    // If a pixel is white but has black pixels above AND below it, 
    // it's a "gap" in a vertical line (common in the number 0). We fill it.
    Bitmap final1x = new Bitmap(baseLayer);
    for (int y = 1; y < baseLayer.Height - 1; y++)
    {
        for (int x = 1; x < baseLayer.Width - 1; x++)
        {
            if (baseLayer.GetPixel(x, y).R == 255) // If White
            {
                // Check for vertical or horizontal gaps
                bool verticalGap = (baseLayer.GetPixel(x, y - 1).R == 0 && baseLayer.GetPixel(x, y + 1).R == 0);
                bool horizontalGap = (baseLayer.GetPixel(x - 1, y).R == 0 && baseLayer.GetPixel(x + 1, y).R == 0);

                if (verticalGap || horizontalGap)
                {
                    final1x.SetPixel(x, y, Color.Black);
                }
            }
        }
    }

    // STEP 3: Professional Scaling (2x)
    int scale = 2;
    int padding = 20;
    Bitmap finalBmp = new Bitmap((original.Width * scale) + (padding * 2), (original.Height * scale) + (padding * 2));

    using (Graphics g = Graphics.FromImage(finalBmp))
    {
        g.Clear(Color.White); // SECOND SAFETY CLEAR
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
        g.DrawImage(final1x, padding, padding, original.Width * scale, original.Height * scale);
    }

    return finalBmp;
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