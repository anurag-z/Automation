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
    // --- FINAL LOGIC: FILTER FIRST, SCALE LATER ---
// --- FINAL FIX: FILTER FIRST (Threshold 40) -> THEN SCALE 3X ---
// --- FINAL ROBUST METHOD (Fixes Black Screen & Thin Text) ---
static Bitmap FilterBlueScreen(Bitmap original)
{
    // 1. Scale Up (3x)
    // We scale FIRST to make the text big (easy for Tesseract).
    // Scaling after filtering can sometimes cause "blobs".
    int scale = 3;
    int padding = 20;
    int w = original.Width * scale;
    int h = original.Height * scale;

    Bitmap newBmp = new Bitmap(w + (padding * 2), h + (padding * 2));

    using (Graphics g = Graphics.FromImage(newBmp))
    {
        // FIX 1: Prevent "Black Screen"
        // We explicitly paint the whole image White first.
        // If we don't do this, the background is Transparent (which looks Black).
        g.Clear(Color.White);

        // Nearest Neighbor is CRITICAL.
        // It keeps the pixels square and sharp.
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
        
        g.DrawImage(original, padding, padding, w, h);
    }

    // 2. The Green Filter (Threshold 60)
    for (int y = 0; y < newBmp.Height; y++)
    {
        for (int x = 0; x < newBmp.Width; x++)
        {
            Color c = newBmp.GetPixel(x, y);

            // LOGIC:
            // Background (Blue) -> Green is 0.
            // Text (Cyan/White) -> Green is 170+.
            // Faint Edges       -> Green is ~80.
            
            // FIX 2: Threshold 60
            // We set this low enough to catch the faint edges of the "0".
            // Since the background is 0, this is perfectly safe.
            
            if (c.G > 60) 
            {
                newBmp.SetPixel(x, y, Color.Black); // Ink (Keep Text)
            }
            else
            {
                newBmp.SetPixel(x, y, Color.White); // Paper (Remove Background)
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
