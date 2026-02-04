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
            // 1. START PROCESS
            ProcessStartInfo processInfo = new ProcessStartInfo();
            processInfo.FileName = "cmd.exe";
            processInfo.WorkingDirectory = @"c:\10405";
            processInfo.Arguments = @"/k fads"; 
            processInfo.UseShellExecute = true;

            Console.WriteLine("Launching Application...");
            Process p = Process.Start(processInfo);
            
            // INCREASED WAIT TIME to ensure window is fully visible
            Thread.Sleep(3000); 

            Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
            Thread.Sleep(500);
            
            Console.WriteLine("Sending F3...");
            PressKey(SC_F3);
            Thread.Sleep(2000);

            // 2. CAPTURE
            Console.WriteLine("Taking screenshot...");
            using (Bitmap original = CaptureWindow(p.MainWindowHandle))
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                // --- DEBUG STEP 1: SAVE RAW IMAGE ---
                // If this image is BLACK, the capture code is failing (Window hidden/minimized).
                string rawPath = Path.Combine(desktop, "debug_1_raw.png");
                original.Save(rawPath, ImageFormat.Png);
                Console.WriteLine($"[CHECK THIS] Saved Raw Screenshot to: {rawPath}");

                // 3. PROCESS
                Console.WriteLine("Processing Image...");
                using (Bitmap processed = FilterSafeMode(original))
                {
                    // --- DEBUG STEP 2: SAVE PROCESSED IMAGE ---
                    // If Raw is good but this is BLACK, the filter is failing.
                    string procPath = Path.Combine(desktop, "debug_2_processed.png");
                    processed.Save(procPath, ImageFormat.Png);
                    Console.WriteLine($"[CHECK THIS] Saved Processed Image to: {procPath}");

                    // 4. OCR
                    string tempFile = "temp_ocr.tif";
                    processed.Save(tempFile, ImageFormat.Tiff);

                    try
                    {
                        using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                        {
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
            Console.WriteLine("CRITICAL ERROR: " + ex.Message);
            Console.WriteLine("Stack Trace: " + ex.StackTrace);
        }
    }

    // --- SAFE MODE FILTER ---
    // This is the simplest, most fail-safe filter possible.
    static Bitmap FilterSafeMode(Bitmap original)
    {
        // 1. Initialize with White Background (Prevents Transparent/Black output)
        Bitmap clean = new Bitmap(original.Width, original.Height);
        using (Graphics g = Graphics.FromImage(clean)) 
        {
            g.Clear(Color.White); 
        }

        for (int y = 0; y < original.Height; y++) 
        {
            for (int x = 0; x < original.Width; x++)
            {
                Color c = original.GetPixel(x, y);
                
                // Brightness Logic
                int b = (int)((c.R * 0.3) + (c.G * 0.59) + (c.B * 0.11));

                // Threshold 50
                // IF Brightness > 50 (Text), DRAW BLACK.
                // IF Brightness < 50 (Background), DO NOTHING (It stays White).
                if (b > 50) 
                {
                    clean.SetPixel(x, y, Color.Black);
                }
            }
        }

        // 2. Scale Up 2x (Simple scale, no filtering)
        int scale = 2;
        int w = original.Width * scale;
        int h = original.Height * scale;
        Bitmap finalBmp = new Bitmap(w, h);

        using (Graphics g = Graphics.FromImage(finalBmp))
        {
            g.Clear(Color.White);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
            g.DrawImage(clean, 0, 0, w, h);
        }
        
        return finalBmp;
    }

    static Bitmap CaptureWindow(IntPtr handle)
    {
        RECT rect;
        GetWindowRect(handle, out rect);

        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;

        Console.WriteLine($"Window Detected Size: {width}x{height} (X:{rect.Left} Y:{rect.Top})");

        if (width <= 0 || height <= 0)
        {
            throw new Exception("Window size is 0x0! The window might be minimized or invalid.");
        }

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