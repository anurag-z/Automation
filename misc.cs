using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using Tesseract;

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
            Bitmap original = CaptureWindow(p.MainWindowHandle);
            
            // 2. PROCESS (The Fix for Blue Screens)
            Console.WriteLine("Applying Blue-Screen Filter...");
            Bitmap processed = FilterBlueScreen(original);

            // Save this image! Open it to verify it looks like a clean Fax document.
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string debugPath = Path.Combine(desktop, "debug_processed_blue.png");
            processed.Save(debugPath, ImageFormat.Png);
            Console.WriteLine($"Debug image saved to: {debugPath}");

            // 3. READ TEXT
            Console.WriteLine("Reading text...");
            
            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            {
                // Force Tesseract to treat the page as a single block of text (Works well for menus)
                using (var page = engine.Process(processed, PageSegMode.SingleBlock))
                {
                    string text = page.GetText();
                    
                    Console.WriteLine("--- FOUND TEXT ---");
                    Console.WriteLine(text);
                    Console.WriteLine("------------------");

                    // Normalize text (Upper case, remove spaces) to match "F7" or "F 7"
                    string cleanText = text.ToUpper().Replace(" ", "");

                    if (cleanText.Contains("F7"))
                    {
                        Console.WriteLine("[PASS] 'F7' instruction found.");
                    }
                    else if (cleanText.Contains("FEDERALFORMS"))
                    {
                        Console.WriteLine("[PASS] Header found (Alternative Pass).");
                    }
                    else
                    {
                        Console.WriteLine("[FAIL] Target text not found.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // --- THE MAGIC FILTER ---
    static Bitmap FilterBlueScreen(Bitmap original)
    {
        // 1. Scale Up (2x) to make pixelated fonts readable
        int scale = 2;
        Bitmap newBmp = new Bitmap(original.Width * scale, original.Height * scale);

        using (Graphics g = Graphics.FromImage(newBmp))
        {
            // NearestNeighbor keeps the DOS font sharp
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
            g.DrawImage(original, 0, 0, newBmp.Width, newBmp.Height);
        }

        // 2. Threshold Loop
        // We look at every pixel. If it's bright (Text), make it BLACK. If it's dark (Blue), make it WHITE.
        for (int y = 0; y < newBmp.Height; y++)
        {
            for (int x = 0; x < newBmp.Width; x++)
            {
                Color c = newBmp.GetPixel(x, y);

                // Calculate "Brightness" (Luminance)
                // White/Cyan text will have high brightness (150-255)
                // Blue background will have low brightness (20-80)
                int brightness = (int)((c.R * 0.3) + (c.G * 0.59) + (c.B * 0.11));

                if (brightness > 100) 
                {
                    newBmp.SetPixel(x, y, Color.Black); // Text -> Black Ink
                }
                else
                {
                    newBmp.SetPixel(x, y, Color.White); // Background -> White Paper
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
