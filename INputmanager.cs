using System;
using System.Runtime.InteropServices;
using System.Threading;

public static class InputManager
{
    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    private const int KEY_UP = 0x0002;
    private const int KEY_SCANCODE = 0x0008;

    // Common Scan Codes for DOS
    public const byte SC_ESCAPE = 0x01;
    public const byte SC_ENTER = 0x1C;

    public static void PressKey(byte scanCode, int delayMs = 150)
    {
        // Key Down
        keybd_event(0, scanCode, KEY_SCANCODE, UIntPtr.Zero);
        Thread.Sleep(50); // Holding time
        
        // Key Up
        keybd_event(0, scanCode, KEY_SCANCODE | KEY_UP, UIntPtr.Zero);
        
        // VDI Buffer Delay: Gives the remote screen time to update
        Thread.Sleep(delayMs); 
    }

    public static void TypeString(string text)
    {
        foreach (char c in text)
        {
            byte code = GetScanCode(c);
            if (code != 0) PressKey(code, 80); // Quick typing delay
        }
    }

    private static byte GetScanCode(char c)
    {
        return char.ToUpper(c) switch
        {
            'A' => 0x1E, 'B' => 0x30, 'C' => 0x2E, 'D' => 0x20, 'E' => 0x12,
            'F' => 0x21, 'G' => 0x22, 'H' => 0x23, 'I' => 0x17, 'J' => 0x24,
            'K' => 0x25, 'L' => 0x26, 'M' => 0x32, 'N' => 0x31, 'O' => 0x18,
            'P' => 0x19, 'Q' => 0x10, 'R' => 0x13, 'S' => 0x1F, 'T' => 0x14,
            'U' => 0x16, 'V' => 0x2F, 'W' => 0x11, 'X' => 0x2D, 'Y' => 0x15,
            'Z' => 0x2C, '0' => 0x0B, '1' => 0x02, '2' => 0x03, '3' => 0x04,
            '4' => 0x05, '5' => 0x06, '6' => 0x07, '7' => 0x08, '8' => 0x09,
            '9' => 0x0A, _ => 0
        };
    }
}