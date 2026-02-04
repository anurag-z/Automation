using System;
using System.Diagnostics;
using System.Drawing; // System.Drawing.Common
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO; // Required for File handling
using Tesseract; // Required for Tesseract

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
        Thread.Sleep(3000); 

        try
        {
            Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
            Thread.Sleep(500);
            
            Console.WriteLine("Sending F3...");
            PressKey(SC_F3);
            Thread.Sleep(2000);

            // 1. CAPTURE (Original size, no scaling)
            Console.WriteLine("Taking screenshot...");
            Bitmap original = CaptureWindow(p.MainWindowHandle);
            
            // Save for your reference
            original.Save("debug_screenshot.png", ImageFormat.Png);

            // 2. READ TEXT (Simple Method)
            Console.WriteLine("Reading text...");
            
            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            {
                // We use the Temp File method because it never fails
                string tempFile = "temp_ocr.tif";
                original.Save(tempFile, ImageFormat.Tiff);

                using (var img = Pix.LoadFromFile(tempFile)) 
                {
                    // SparseText worked best for your menu layout
                    using (var page = engine.Process(img, PageSegMode.SparseText))
                    {
                        string text = page.GetText();
                        
                        Console.WriteLine("--- FOUND TEXT ---");
                        Console.WriteLine(text);
                        Console.WriteLine("------------------");

                        // --- THE FIX: IGNORE SPACES ---
                        // Tesseract often sees "F3" as "F 3". We remove spaces to catch it.
                        string cleanText = text.Replace(" ", "").ToUpper();

                        if (cleanText.Contains("F3") || cleanText.Contains("F03"))
                        {
                            Console.WriteLine("[PASS] F3 Found.");
                        }
                        else if (text.Contains("FADS"))
                        {
                            Console.WriteLine("[PASS] 'FADS' header found (F3 Key likely worked).");
                        }
                        else
                        {
                            Console.WriteLine("[FAIL] Could not verify screen.");
                        }
                    }
                }
                
                // Cleanup
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // --- HELPER ---
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
