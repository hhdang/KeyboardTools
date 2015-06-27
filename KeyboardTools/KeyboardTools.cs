using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Threading;

namespace KeyboardTools
{
    class KeyboardTools
    {
        [DllImport("user32.dll")]
        public static extern bool keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);

        [DllImport("user32.dll")]
        static extern byte VkKeyScan(char ch);

        const uint WM_CHAR = 0x102;
        const uint WM_PASTE = 0x0302;
        const int VK_UP = 0x26;
        const int VK_DOWN = 0x28;
        const int VK_LEFT = 0x25;
        const int VK_RIGHT = 0x27;
        const uint KEYEVENTF_KEYUP = 0x0002;
        const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

        public static void typeText(IntPtr hWnd, String myString)
        {
            char c;
            int sleepTime;
            Random random = new Random();

            // For each character in the string
            for (int i = 0; i < myString.Length; i++)
            {
                c = (char)myString[i];
                // Press Key
                sleepTime = random.Next(20) * 10 + 100;
                pressKey(hWnd, c);
                Thread.Sleep(sleepTime);

                // Release Key
                sleepTime = random.Next(20) * 10 + 100;
                releaseKey(hWnd, c);
                Thread.Sleep(sleepTime);
            }
            char enter = (char)0x1C;

            // Press Enter / Return
            sleepTime = random.Next(20) * 10 + 100;
            pressKey(hWnd, enter);
            Thread.Sleep(sleepTime);

            // Release Return;
            sleepTime = random.Next(20) * 10 + 100;
            releaseKey(hWnd, enter);
            Thread.Sleep(sleepTime);
        }

        public static void pressKey(IntPtr hWnd, char key)
        {
            if (hWnd == IntPtr.Zero)
            {
                throw new Exception("Application is not running.");
            }
            if ((scanKeyDictionary.ContainsKey(key)) && (virtualKeyDictionary.ContainsKey(key)))
            {
                keybd_event(virtualKeyDictionary[key], scanKeyDictionary[key], 0, 0);
            }
        }

        public static void releaseKey(IntPtr hWnd, char key)
        {
            if (hWnd == IntPtr.Zero)
            {
                throw new Exception("Application is not running.");
            }
            if ((scanKeyDictionary.ContainsKey(key)) && (virtualKeyDictionary.ContainsKey(key)))
            {
                keybd_event(virtualKeyDictionary[key], scanKeyDictionary[key], KEYEVENTF_KEYUP, 0);
            }
        }

        public static byte getScanKeyDictionary(char key)
        {
            return scanKeyDictionary[key];
        }

        /// <summary>
        /// Collects the system Virtual Key Codes as returned from the Users32.DLL VkKeyScan function from 0x01 to 0xFF.
        /// VkKeyScan('0') broke this function.  Since this value is not used it was ignored.
        /// </summary>
        /// <returns>
        /// String, complete with tabs and line feeds, for display in a TextBox control.
        /// </returns>
        public static String showSystemVirtualByteCodes()
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 1; i < 256; i++)
            {
                sb.Append(i.ToString());
                sb.Append(":\t");
                sb.Append((char)i);
                sb.Append("\t");
                sb.Append(i.ToString("X"));
                sb.Append("\t");
                byte result = VkKeyScan((char)(i));
                sb.Append(result.ToString("X"));
                sb.Append("\r\n");

            }

            return sb.ToString();
        }

        #region ScanKey Dictionary
        // There is an ASCII code, a UNICODE, a ScanKey code, and a Virtual Key code

        // US Standard Keyboard Layout
        private static Dictionary<char, byte> scanKeyDictionary = new Dictionary<char, byte>()
        {
            {' ', 0x39},            // Spacebar
            {(char)0x1C, 0x1C},     // Enter or Return
            {'/', 0x35},            // Forward Slash
            // Standard Alpha-Numeric
            {'1', 0x02}, 
            {'2', 0x03}, 
            {'3', 0x04},
            {'4', 0x05}, 
            {'5', 0x06}, 
            {'6', 0x07}, 
            {'7', 0x08}, 
            {'8', 0x09}, 
            {'9', 0x0A}, 
            {'0', 0x0B},
            {'a', 0x1E}, 
            {'b', 0x30}, 
            {'c', 0x2E}, 
            {'d', 0x20}, 
            {'e', 0x12}, 
            {'f', 0x21},
            {'g', 0x22}, 
            {'h', 0x23}, 
            {'i', 0x17}, 
            {'j', 0x24},
            {'k', 0x25}, 
            {'l', 0x26}, 
            {'m', 0x32}, 
            {'n', 0x31}, 
            {'o', 0x18}, 
            {'p', 0x19}, 
            {'q', 0x10}, 
            {'r', 0x13}, 
            {'s', 0x1F}, 
            {'t', 0x14},
            {'u', 0x16}, 
            {'v', 0x2F}, 
            {'w', 0x11}, 
            {'x', 0x2D}, 
            {'y', 0x15}, 
            {'z', 0x2C}
        };
        #endregion

        #region Virtual Key Dictionary
        // There is an ASCII code, a UNICODE, a ScanKey code, and a Virtual Key code
        // US Standard Keyboard Layout
        private static Dictionary<char, byte> virtualKeyDictionary = new Dictionary<char, byte>()
        {
            {' ', 0x20},            // Spacebar
            {(char)0x1C, 0x0D},     // Enter or Return            
            {'/', 0xBF},            // Forward Slash
            // Standard Alpha-Numeric
            {'0', 0x30},
            {'1', 0x31},
            {'2', 0x32},
            {'3', 0x33},
            {'4', 0x34},
            {'5', 0x35},
            {'6', 0x36},
            {'7', 0x37},
            {'8', 0x38},
            {'9', 0x39},
            {'a', 0x41},
            {'b', 0x42},
            {'c', 0x43},
            {'d', 0x44},
            {'e', 0x45},
            {'f', 0x46},
            {'g', 0x47},
            {'h', 0x48},
            {'i', 0x49},
            {'j', 0x4A},
            {'k', 0x4B},
            {'l', 0x4C},
            {'m', 0x4D},
            {'n', 0x4E},
            {'o', 0x4F},
            {'p', 0x50},
            {'q', 0x51},
            {'r', 0x52},
            {'s', 0x53},
            {'t', 0x54},
            {'u', 0x55},
            {'v', 0x56},
            {'w', 0x57},
            {'x', 0x58},
            {'y', 0x59},
            {'z', 0x5A}
        };
        #endregion

    }
}
