using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PostAction
{
    class Program
    {
        [DllImport("USER32.DLL", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint msg, uint wParam, uint lParam);


        [DllImport("USER32.DLL")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr i);

        public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);

        /// <summary>
        /// Callback method to be used when enumerating windows.
        /// </summary>
        /// <param name="handle">Handle of the next window</param>
        /// <param name="pointer">Pointer to a GCHandle that holds a reference to the list to fill</param>
        /// <returns>True to continue the enumeration, false to bail</returns>
        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(pointer);
            List<IntPtr> list = gch.Target as List<IntPtr>;
            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }
            list.Add(handle);
            //  You can modify this to check to see if you want to cancel the operation, then return a null here
            return true;
        }


        public static List<IntPtr> GetChildWindows(IntPtr parent)
        {
            List<IntPtr> result = new List<IntPtr>();
            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result;
        }

        static void Main(string[] args)
        {
            Process[] processes = Process.GetProcessesByName("Revit");
            if (processes.Length > 0 && args.Length>0)
            {
                int timeOut = 10;
                DateTime time;
                int delta;
                IntPtr strPtr;
                IntPtr hWnd = IntPtr.Zero;
                time = DateTime.Now;
                do
                {
                    hWnd = FindWindow(null, "Load File as Group");
                    Thread.Sleep(1);
                    delta = (DateTime.Now - time).Seconds;
                    if(delta>timeOut)
                    {
                        break;
                    }
                } while (hWnd == IntPtr.Zero);
                if (delta <= timeOut)
                {
                    time = DateTime.Now;
                    do
                    {
                        Thread.Sleep(1);
                        hWnd = FindWindow(null, "Load File as Group");
                        delta = (DateTime.Now - time).Seconds;
                        if (delta > timeOut || hWnd==IntPtr.Zero)
                        {
                            break;
                        }
                        try
                        {
                            strPtr = Marshal.StringToCoTaskMemAuto(args[0]);
                            SendMessage(GetChildWindows(hWnd)[1], 0X000C, 0, (uint)strPtr);
                            PostMessage(GetChildWindows(hWnd)[5], 0x100, 0xD, 0x1C0001);
                            PostMessage(GetChildWindows(hWnd)[5], 0x101, 0xD, 0xC01C0001);
                        }
                        catch(Exception ex)
                        {
                        }
                    } while (hWnd != IntPtr.Zero );
                }
            }
        }
    }
}
