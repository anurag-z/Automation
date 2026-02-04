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
    static Bitmap FilterBlueScreen(Bitmap original)
    {
        // 1. Scale Up (3x) for Tesseract
        int scale = 3;
        Bitmap newBmp = new Bitmap(original.Width * scale, original.Height * scale);

        using (Graphics g = Graphics.FromImage(newBmp))
        {
            // KEEP IT SHARP! (Nearest Neighbor is mandatory for DOS)
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
            g.DrawImage(original, 0, 0, newBmp.Width, newBmp.Height);
        }

        // 2. The Filter Loop
        for (int y = 0; y < newBmp.Height; y++)
        {
            for (int x = 0; x < newBmp.Width; x++)
            {
                Color c = newBmp.GetPixel(x, y);

                // LOGIC:
                // We sum the Red and Green values.
                // Background (Blue) = 0 Red + 0 Green = 0 Total.
                // Cyan Text = 0 Red + 255 Green = 255 Total.
                // White Text = 255 Red + 255 Green = 510 Total.
                
                // We set the cutoff at 100. 
                // This ignores faint background noise but catches all text.
                int brightnessSum = c.R + c.G;

                if (brightnessSum > 100) 
                {
                    newBmp.SetPixel(x, y, Color.Black); // Ink
                }
                else
                {
                    newBmp.SetPixel(x, y, Color.White); // Paper
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
