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
    // --- KEYBOARD & WINDOW SETUP ---
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
        // 1. START THE APP
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

            Console.WriteLine("Taking screenshot...");
            using (Bitmap original = CaptureWindow(p.MainWindowHandle))
            {
                // 2. PROCESS IMAGE (Green > 70)
                Console.WriteLine("Processing image...");
                using (Bitmap processed = FilterBlueScreen(original))
                {
                    // Save to desktop so you can see the result
                    string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    processed.Save(Path.Combine(desktop, "debug_final_v3.png"), ImageFormat.Png);

                    string tempFile = "temp_ocr.tif";
                    processed.Save(tempFile, ImageFormat.Tiff);

                    try
                    {
                        using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                        {
                            // --- CRITICAL FIX: THE WHITELIST ---
                            // This stops "1040" from reading as "1242" or "1O4O".
                            // It forces Tesseract to pick the best matching NUMBER or UPPERCASE LETTER.
                            engine.SetVariable("tessedit_char_whitelist", "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ,.-() ");
                            
                            using (var img = Pix.LoadFromFile(tempFile)) 
                            {
                                // SingleBlock mode is best for menus
                                using (var page = engine.Process(img, PageSegMode.SingleBlock))
                                {
                                    string text = page.GetText();
                                    Console.WriteLine("\n--- OCR OUTPUT ---");
                                    Console.WriteLine(text);
                                    Console.WriteLine("------------------");

                                    string cleanText = text.ToUpper().Replace(" ", "");
                                    
                                    // 3. VERIFICATION
                                    if (cleanText.Contains("F7"))
                                        Console.WriteLine("[PASS] 'F7' Found.");
                                    else if (cleanText.Contains("1040"))
                                        Console.WriteLine("[PASS] '1040' Found.");
                                    else
                                        Console.WriteLine("[FAIL] Text not matched.");
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

    // --- FINAL FILTER METHOD (Green > 70) ---
    static Bitmap FilterBlueScreen(Bitmap original)
    {
        // 1. Scale Up (2x)
        int scale = 2;
        int padding = 20;
        int w = original.Width * scale;
        int h = original.Height * scale;

        Bitmap newBmp = new Bitmap(w + (padding * 2), h + (padding * 2));

        using (Graphics g = Graphics.FromImage(newBmp))
        {
            g.Clear(Color.White); 
            
            // NEAREST NEIGHBOR: Keeps pixels square and sharp (No Blur!)
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
            
            g.DrawImage(original, padding, padding, w, h);
        }

        // 2. The Filter
        for (int y = 0; y < newBmp.Height; y++)
        {
            for (int x = 0; x < newBmp.Width; x++)
            {
                Color c = newBmp.GetPixel(x, y);

                // LOGIC:
                // We lowered the threshold to 70.
                // This captures the edges of the letters that were fading out before.
                // Since the Background is 0 Green, it stays White.
                
                if (c.G > 70) 
                {
                    newBmp.SetPixel(x, y, Color.Black); // Keep Text
                }
                else
                {
                    newBmp.SetPixel(x, y, Color.White); // Remove Background
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
