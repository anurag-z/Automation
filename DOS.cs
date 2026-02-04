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
    }
}
