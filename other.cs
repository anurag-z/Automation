using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;

class ExternalReader
{
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool AttachConsole(int dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool FreeConsole();

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool ReadConsoleOutputCharacter(IntPtr hConsoleOutput, [Out] StringBuilder lpCharacter, uint nLength, COORD dwReadCoord, out uint lpNumberOfCharsRead);

    [StructLayout(LayoutKind.Sequential)]
    public struct COORD { public short X; public short Y; }

    private const int STD_OUTPUT_HANDLE = -11;

    public static void Main()
    {
        // 1. Find the FADS process
        var process = Process.GetProcessesByName("FADS").FirstOrDefault();
        if (process == null) {
            Console.WriteLine("FADS process not found.");
            return;
        }

        // 2. Attach to the FADS console buffer
        if (AttachConsole(process.Id))
        {
            IntPtr hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
            
            // 3. Read Row 23 (The "Continued..." line)
            COORD readCoord = new COORD { X = 0, Y = 23 };
            StringBuilder sb = new StringBuilder(80);
            uint charsRead;

            if (ReadConsoleOutputCharacter(hConsole, sb, 80, readCoord, out charsRead))
            {
                string line = sb.ToString();
                string status = line.Substring(30, 20).Trim(); // "Continued..."
                string branch = line.Substring(70, 10).Trim(); // "BRCH"

                Console.WriteLine($"Row 23 Text: {line}");
                Console.WriteLine($"Extracted: {status} | {branch}");
            }

            // 4. Always detach when done
            FreeConsole();
        }
        else
        {
            Console.WriteLine("Could not attach. Ensure FADS is a console-based app.");
        }
    }
}