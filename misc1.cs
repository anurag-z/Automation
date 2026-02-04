using System;
using System.Drawing; // Reference: System.Drawing.Common
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using Tesseract; // Reference: Tesseract NuGet

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
            // 1. LAUNCH & SETUP
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
                // 2. PROCESS IMAGE
                Console.WriteLine("Processing...");
                using (Bitmap processed = FilterGrayscaleBold(original))
                {
                    // DEBUG: Save to desktop to VERIFY the image is white with black text
                    string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    processed.Save(Path.Combine(desktop, "final_debug.png"), ImageFormat.Png);

                    // 3. OCR
                    string tempFile = "temp_ocr.tif";
                    processed.Save(tempFile, ImageFormat.Tiff);

                    try
                    {
                        using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                        {
                            // CRITICAL: Allow Numbers, Uppercase, and Punctuation
                            engine.SetVariable("tessedit_char_whitelist", "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ,.-() ");
                            
                            using (var img = Pix.LoadFromFile(tempFile)) 
                            {
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
            Console.WriteLine("ERROR: " + ex.Message);
        }
    }

    // --- THE FIX ---
    static Bitmap FilterGrayscaleBold(Bitmap original)
    {
       // STEP 1: Clean the image at original size (1x)
    // We simply separate Light Text from Dark Background.
    Bitmap clean1x = new Bitmap(original.Width, original.Height);
    
    // Safety: Start with a White canvas
    using (Graphics g = Graphics.FromImage(clean1x)) { g.Clear(Color.White); }

    for (int y = 0; y < original.Height; y++) 
    {
        for (int x = 0; x < original.Width; x++)
        {
            Color c = original.GetPixel(x, y);

            // Calculate Brightness (0 to 255)
            // This works for Cyan, White, Yellow - any bright text color.
            int brightness = (int)((c.R * 0.3) + (c.G * 0.59) + (c.B * 0.11));

            // Threshold: 60
            // Background is usually ~20. Text is usually > 150.
            // 60 is the perfect cutoff to remove the background noise.
            if (brightness > 60) 
            {
                clean1x.SetPixel(x, y, Color.Black); // Ink
            }
            // We don't need "else { SetWhite }" because we cleared the canvas to White at the start.
        }
    }

    // STEP 2: Scale Up (2x)
    // We scale up to make it easier for Tesseract to read the clean letters.
    int scale = 2;
    int padding = 20; // Add border
    int w = original.Width * scale;
    int h = original.Height * scale;

    Bitmap finalBmp = new Bitmap(w + (padding * 2), h + (padding * 2));

    using (Graphics g = Graphics.FromImage(finalBmp))
    {
        g.Clear(Color.White); // White background

        // CRITICAL: NearestNeighbor
        // Since we are using a computer font (Consolas), we want exact pixels.
        // We do NOT want the "smoothing" that makes text blurry.
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
        
        g.DrawImage(clean1x, padding, padding, w, h);
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
