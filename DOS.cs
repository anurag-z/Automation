using System;
using System.Diagnostics;
using System.Runtime.InteropServices; 
using System.Threading;
using Microsoft.VisualBasic;

class Program
{
    [DllImport("user32.dll")]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    // Constants
    const int VK_F3 = 0x72; 
    const byte SC_F3 = 0x3D; // <--- The Hardware Scan Code for F3
    
    // Flags
    const uint KEYEVENTF_KEYUP = 0x0002;
    const uint KEYEVENTF_SCANCODE = 0x0008; // <--- Tells windows to use the Scan Code

    static void Main()
    {
        ProcessStartInfo processInfo = new ProcessStartInfo();
        processInfo.FileName = "cmd.exe";
        processInfo.WorkingDirectory = @"c:\10405";
        processInfo.Arguments = @"/k fads"; 
        processInfo.UseShellExecute = true;
    //processInfo.Arguments = @"/k mode con: cols=80 lines=50 && fads";
        Process p = Process.Start(processInfo);

        Thread.Sleep(3000); 

        try
        {
            // Focus the window
            Interaction.AppActivate(p.Id);
            Thread.Sleep(500);

            // 1. Press F3 DOWN using Scan Code
            // We pass '0' for the first argument because we are using the ScanCode flag
            keybd_event(0, SC_F3, KEYEVENTF_SCANCODE, UIntPtr.Zero);
            
            Thread.Sleep(100); 
            
            // 2. Release F3 UP
            keybd_event(0, SC_F3, KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP, UIntPtr.Zero);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        try {
    // 1. Send the keys to open the menu and copy text
    // Alt+Space opens window menu, 'E' for Edit, 'S' for Select All
    SendKeys.SendWait("%{SPACE}"); // Alt + Space
    Thread.Sleep(500);
    SendKeys.SendWait("es");       // Edit -> Select All
    
    Thread.Sleep(500);
    SendKeys.SendWait("{ENTER}");  // Copy to clipboard
    
    // 2. Get the text from the Clipboard
    string fullScreenText = Clipboard.GetText();

    // 3. Split the text into separate lines
    string[] lines = fullScreenText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

    // READ THE TOP LINE (Header)
    if (lines.Length > 0)
    {
        string topLine = lines[0];
        Console.WriteLine("TOP OF SCREEN: " + topLine);
    }

    // READ THE BOTTOM LINE (Status Bar)
    // We check for empty lines at the end, sometimes copy adds whitespace
    if (lines.Length > 1)
    {
        // Often the last line is empty, so we take the one before it
        string bottomLine = lines[lines.Length - 2]; 
        Console.WriteLine("BOTTOM OF SCREEN: " + bottomLine);
    }
}
catch (Exception ex)
{
    Console.WriteLine("Could not read screen: " + ex.Message);
}
    }
}
