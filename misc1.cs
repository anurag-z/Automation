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
    // ---------------- WIN32 ----------------
    [DllImport("user32.dll")]
    static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, int nFlags);

    [DllImport("user32.dll")]
    static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    const int KEYEVENTF_KEYUP = 0x0002;
    const int KEYEVENTF_SCANCODE = 0x0008;
    const byte SC_F3 = 0x3D;

    static void Main()
    {
        try
        {
            // 1️⃣ LAUNCH DOS APP
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                WorkingDirectory = @"c:\10405",
                Arguments = "/k fads",
                UseShellExecute = true
            };

            Process p = Process.Start(psi);
            p.WaitForInputIdle();
            Thread.Sleep(1000);

            // Bring to foreground
            Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
            Thread.Sleep(500);

            // Example key
            PressKey(SC_F3);
            Thread.Sleep(1000);

            // 2️⃣ CAPTURE DOS CLIENT AREA ONLY
            using (Bitmap captured = CaptureDosClient(p.MainWindowHandle))
            {
                // 3️⃣ FILTER FOR OCR (TUNED FOR YOUR SCREEN)
                using (Bitmap processed = FilterDosBlueScreen(captured))
                {
                    // DEBUG: Save image to verify
                    string debugPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        "dos_ocr_debug.png");

                    processed.Save(debugPath, ImageFormat.Png);
                    Console.WriteLine("Saved OCR image: " + debugPath);

                    // 4️⃣ OCR
                    using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                    {
                        engine.DefaultPageSegMode = PageSegMode.SingleColumn;

                        engine.SetVariable("preserve_interword_spaces", "1");
                        engine.SetVariable("textord_force_make_prop_words", "F");

                        engine.SetVariable(
                            "tessedit_char_whitelist",
                            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-.,() ");

                        using (var pix = PixConverter.ToPix(processed))
                        using (var page = engine.Process(pix))
                        {
                            Console.WriteLine("\n--- OCR OUTPUT ---\n");
                            Console.WriteLine(page.GetText());
                            Console.WriteLine("------------------");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR: " + ex.Message);
        }
    }

    // ---------------- CAPTURE ----------------
    static Bitmap CaptureDosClient(IntPtr hWnd)
    {
        GetClientRect(hWnd, out RECT rc);

        int width = rc.Right - rc.Left;
        int height = rc.Bottom - rc.Top;

        Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);

        using (Graphics g = Graphics.FromImage(bmp))
        {
            IntPtr hdc = g.GetHdc();
            try
            {
                // PW_CLIENTONLY = 1 → no borders/title
                PrintWindow(hWnd, hdc, 1);
            }
            finally
            {
                g.ReleaseHdc(hdc);
            }
        }

        return bmp;
    }

    // ---------------- IMAGE FILTER (DOS BLUE UI) ----------------
    static Bitmap FilterDosBlueScreen(Bitmap original)
    {
        Bitmap output = new Bitmap(original.Width, original.Height);

        for (int y = 0; y < original.Height; y++)
        {
            for (int x = 0; x < original.Width; x++)
            {
                Color c = original.GetPixel(x, y);

                // Dark blue text
                bool isDarkBlueText =
                    c.B > c.R + 20 &&
                    c.B > c.G + 20 &&
                    c.B < 200;

                // White text
                bool isWhiteText =
                    c.R > 220 && c.G > 220 && c.B > 220;

                if (isDarkBlueText || isWhiteText)
                    output.SetPixel(x, y, Color.Black);
                else
                    output.SetPixel(x, y, Color.White);
            }
        }

        // Scale 2x (nearest neighbor)
        Bitmap scaled = new Bitmap(output.Width * 2, output.Height * 2);
        using (Graphics g = Graphics.FromImage(scaled))
        {
            g.Clear(Color.White);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.DrawImage(output, 0, 0, scaled.Width, scaled.Height);
        }

        scaled.SetResolution(300, 300);
        return scaled;
    }

    // ---------------- KEY PRESS ----------------
    static void PressKey(byte scanCode)
    {
        keybd_event(0, scanCode, KEYEVENTF_SCANCODE, UIntPtr.Zero);
        Thread.Sleep(100);
        keybd_event(0, scanCode, KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
