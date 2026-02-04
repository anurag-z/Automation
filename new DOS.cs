using System;
using System.Diagnostics;
using System.Drawing; // Requires System.Drawing.Common
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms; // Requires System.Windows.Forms
using Tesseract; // Requires 'Tesseract' NuGet package

class Program
{
    // --- KEYBOARD & WINDOW SETUP ---
    [DllImport("user32.dll")]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    const int KEYEVENTF_KEYUP = 0x0002;
    const int KEYEVENTF_SCANCODE = 0x0008;
    const byte SC_F3 = 0x3D;

    static void Main()
    {
        // 1. Launch App
        ProcessStartInfo processInfo = new ProcessStartInfo();
        processInfo.FileName = "cmd.exe";
        processInfo.WorkingDirectory = @"c:\10405";
        processInfo.Arguments = @"/k fads";
        processInfo.UseShellExecute = true;

        Process p = Process.Start(processInfo);
        Thread.Sleep(3000); // Wait for launch

        try
        {
            // 2. Focus & Action
            Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
            Thread.Sleep(500);
            
            Console.WriteLine("Sending F3...");
            PressKey(SC_F3);
            Thread.Sleep(2000); // Wait for screen to update

            // 3. CAPTURE SCREENSHOT
            Console.WriteLine("Taking screenshot...");
            Bitmap screenshot = CaptureWindow(p.MainWindowHandle);
            
            // Save it just so you can verify it looks right
            string imagePath = "debug_screenshot.png";
            screenshot.Save(imagePath, ImageFormat.Png);

            // 4. READ TEXT (OCR)
            Console.WriteLine("Reading text from image...");
            
            // Point this to where you put the 'tessdata' folder
            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(imagePath))
                {
                    using (var page = engine.Process(img))
                    {
                        string text = page.GetText();
                        
                        Console.WriteLine("--- OCR RESULT ---");
                        
                        // Verification Logic
                        if (text.Contains("FADS"))
                        {
                            Console.WriteLine("[PASS] 'FADS' found in screenshot.");
                        }
                        else
                        {
                            Console.WriteLine("[FAIL] Could not find expected text.");
                            Console.WriteLine("Found this instead: \n" + text.Substring(0, Math.Min(100, text.Length)));
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // --- HELPER: Screenshot Active Window ---
    static Bitmap CaptureWindow(IntPtr handle)
    {
        // Get the size of the window
        RECT rect;
        GetWindowRect(handle, out rect);
        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;

        // Create a bitmap of that size
        Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        
        // Draw the screen into the bitmap
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(rect.Left, rect.Top, 0, 0, bmp.Size, CopyPixelOperation.SourceCopy);
        }
        return bmp;
    }
    //https://github.com/tesseract-ocr/tessdata/blob/main/eng.traineddata

    static void PressKey(byte scanCode)
    {
        keybd_event(0, scanCode, KEYEVENTF_SCANCODE, UIntPtr.Zero);
        Thread.Sleep(100);
        keybd_event(0, scanCode, KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
