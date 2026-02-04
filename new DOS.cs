using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.VisualBasic;

class Program
{
    // --- NATIVE METHODS ---
    [DllImport("user32.dll")]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool AttachConsole(uint dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool FreeConsole();

    [DllImport("kernel32.dll")]
    static extern bool AllocConsole();

    [DllImport("kernel32.dll")]
    static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll")]
    static extern bool GetConsoleScreenBufferInfo(IntPtr hConsoleOutput, out CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    static extern bool ReadConsoleOutputCharacter(IntPtr hConsoleOutput, [Out] StringBuilder lpCharacter, uint nLength, Coord dwReadCoord, out uint lpNumberOfCharsRead);

    // --- CONSTANTS ---
    const int STD_OUTPUT_HANDLE = -11;
    const int KEYEVENTF_KEYUP = 0x0002;
    const int KEYEVENTF_SCANCODE = 0x0008;
    const byte SC_F3 = 0x3D;

    // --- STRUCTS ---
    [StructLayout(LayoutKind.Sequential)]
    public struct Coord
    {
        public short X;
        public short Y;
        public Coord(short x, short y) { X = x; Y = y; }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SMALL_RECT
    {
        public short Left, Top, Right, Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CONSOLE_SCREEN_BUFFER_INFO
    {
        public Coord dwSize;
        public Coord dwCursorPosition;
        public ushort wAttributes;
        public SMALL_RECT srWindow;
        public Coord dwMaximumWindowSize;
    }

    // --- MAIN ---
    static void Main()
    {
        // 1. Launch App (Without forcing mode)
        ProcessStartInfo processInfo = new ProcessStartInfo();
        processInfo.FileName = "cmd.exe";
        processInfo.WorkingDirectory = @"c:\10405";
        processInfo.Arguments = @"/k fads"; // No 'mode' command
        processInfo.UseShellExecute = true;

        Process p = Process.Start(processInfo);
        Thread.Sleep(3000); 

        try
        {
            // 2. Interact
            Interaction.AppActivate(p.Id);
            Thread.Sleep(500);
            PressKey(SC_F3);
            Thread.Sleep(2000);

            // 3. READ SCREEN (Returns a list of lines)
            List<string> screenLines = ReadConsoleLines(p.Id);

            // 4. RE-OPEN CONSOLE TO SHOW RESULTS
            AllocConsole(); 
            Console.WriteLine($"--- Captured {screenLines.Count} Lines ---");

            // --- VERIFICATION EXAMPLES ---
            
            // A. Check the TOP line (Header)
            if (screenLines.Count > 0)
            {
                string header = screenLines[0].Trim();
                Console.WriteLine($"Header: '{header}'");
                
                if (header.Contains("FADS")) 
                    Console.WriteLine("[PASS] Header correct.");
                else 
                    Console.WriteLine("[FAIL] Header mismatch.");
            }

            // B. Check the BOTTOM line (Status Bar)
            // Note: The buffer might be taller than the window, so we check the last non-empty line
            for (int i = screenLines.Count - 1; i >= 0; i--)
            {
                if (!string.IsNullOrWhiteSpace(screenLines[i]))
                {
                    Console.WriteLine($"Status Line Found (Row {i}): '{screenLines[i].Trim()}'");
                    break;
                }
            }

            Console.ReadLine(); // Pause to see result
        }
        catch (Exception ex)
        {
            AllocConsole();
            Console.WriteLine("Error: " + ex.Message);
            Console.ReadLine();
        }
    }

    // --- HELPER: Read Lines Dynamically ---
    static List<string> ReadConsoleLines(int processId)
    {
        List<string> lines = new List<string>();
        FreeConsole(); 

        if (AttachConsole((uint)processId))
        {
            IntPtr stdOut = GetStdHandle(STD_OUTPUT_HANDLE);
            CONSOLE_SCREEN_BUFFER_INFO csbi;

            // 1. ASK the window how big it is
            if (GetConsoleScreenBufferInfo(stdOut, out csbi))
            {
                int width = csbi.dwSize.X;
                int height = csbi.dwSize.Y;
                int totalChars = width * height;

                StringBuilder sb = new StringBuilder(totalChars);
                uint read = 0;

                // 2. Read the whole buffer
                ReadConsoleOutputCharacter(stdOut, sb, (uint)totalChars, new Coord(0, 0), out read);

                string allText = sb.ToString();

                // 3. Slice it into lines based on the DETECTED width
                for (int i = 0; i < height; i++)
                {
                    // Safety check to avoid index errors
                    if (i * width + width <= allText.Length)
                    {
                        string line = allText.Substring(i * width, width);
                        lines.Add(line);
                    }
                }
            }
            FreeConsole();
        }
        return lines;
    }

    static void PressKey(byte scanCode)
    {
        keybd_event(0, scanCode, KEYEVENTF_SCANCODE, UIntPtr.Zero);
        Thread.Sleep(100);
        keybd_event(0, scanCode, KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
