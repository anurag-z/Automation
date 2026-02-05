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

   static Bitmap FilterUniversalOCR(Bitmap original)
{
    // STEP 1: Process at 1x size to keep character geometry intact
    Bitmap bold1x = new Bitmap(original.Width, original.Height);
    
    using (Graphics g = Graphics.FromImage(bold1x)) 
    { 
        g.Clear(Color.White); // Critical: Start with 'Paper'
    }

    for (int y = 0; y < original.Height - 1; y++) 
    {
        for (int x = 0; x < original.Width - 1; x++)
        {
            Color c = original.GetPixel(x, y);

            // Luminance thresholding
            int brightness = (int)((c.R * 0.3) + (c.G * 0.59) + (c.B * 0.11));

            if (brightness > 40) // Captures both Cyan and White text
            {
                // Current pixel
                bold1x.SetPixel(x, y, Color.Black);
                
                // UNIVERSAL BOLDING: This repairs horizontal bars in 'E', 'F' 
                // and vertical loops in '0', '8', and 'S'
                bold1x.SetPixel(x + 1, y, Color.Black); // Expand Right
                bold1x.SetPixel(x, y + 1, Color.Black); // Expand Down
            }
        }
    }

    // STEP 2: Professional Scaling (2x Nearest Neighbor)
    int scale = 2;
    int pad = 20;
    Bitmap finalBmp = new Bitmap((original.Width * scale) + (pad * 2), (original.Height * scale) + (pad * 2));

    using (Graphics g = Graphics.FromImage(finalBmp))
    {
        g.Clear(Color.White); // Prevent transparency errors
        
        // Nearest Neighbor keeps the bolded text sharp for the engine
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
        
        g.DrawImage(bold1x, pad, pad, original.Width * scale, original.Height * scale);
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