using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using SHDocVw;


namespace IEcctv
{
    public static class WINAPI
    {
        public const int SW_MAXIMIZE = 3;

        [DllImport("user32.dll")]
        public static extern int ShowWindow(IntPtr hWnd, int nCmdShow);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        public static extern int GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);
    }

    internal static class Program
    {
        static ManualResetEventSlim doExit = new ManualResetEventSlim(false);
        static bool doNotExit = true;

        static string logPath;
        static string imagePath;

        static IntPtr IE_hwnd;

        static ImageCodecInfo imgCodec;
        static EncoderParameters imgEncoderParams;
        static Bitmap bmp;
        static Graphics bmpGfx;

        private static void doExit_set() {
            doExit.Set();
            doNotExit = false;
        }

        private static void Initialize() {
            logPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + $"\\Microsoft\\Log4";
            imagePath = logPath + $"\\{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}";
            Directory.CreateDirectory(imagePath);

            //Int64 quality = (Int64)Math.Round(10 / (Graphics.FromHwnd(IntPtr.Zero).DpiX / 96));
            Int64 quality = (Int64)Math.Max(Math.Round(960 / Graphics.FromHwnd(IntPtr.Zero).DpiX, MidpointRounding.AwayFromZero), 1);

            imgCodec = ImageCodecInfo.GetImageEncoders().Single(x => x.MimeType == "image/jpeg");
            imgEncoderParams = new EncoderParameters(1);
            imgEncoderParams.Param[0] = new EncoderParameter(Encoder.Quality, quality);
            bmp = new Bitmap(SystemInformation.VirtualScreen.Width, SystemInformation.VirtualScreen.Height, PixelFormat.Format24bppRgb);
            bmpGfx = Graphics.FromImage(bmp);
            bmpGfx.CompositingMode = CompositingMode.SourceCopy;
            bmpGfx.CompositingQuality = CompositingQuality.HighSpeed;
            bmpGfx.InterpolationMode = InterpolationMode.NearestNeighbor;
            bmpGfx.SmoothingMode = SmoothingMode.None;
        }

        private static void KillAllIE() {
            foreach (var item in Process.GetProcessesByName("iexplore")) {
                item.Kill();
            }
        }

        private static void IExistanceCheck() {
            try {
                int ie_process;
                WINAPI.GetWindowThreadProcessId(IE_hwnd, out ie_process);
                Process IEproc = Process.GetProcessById(ie_process);
                IEproc.WaitForExit();
            } finally {
                doExit_set();
            }
        }

        private static void MakeCapture() {
            Stopwatch stw = new Stopwatch();
            WINAPI.RECT rect;
            Size wndsz = new Size();

            Thread.Sleep(1000);
            while (doNotExit) {
                stw.Restart();
                try {
                    WINAPI.GetWindowRect(IE_hwnd, out rect);
                    wndsz.Width = rect.Right - rect.Left;
                    wndsz.Height = rect.Bottom - rect.Top;
                    bmpGfx.Clear(Color.Black);
                    bmpGfx.CopyFromScreen(rect.Left, rect.Top, 0, 0, wndsz, CopyPixelOperation.SourceCopy);

                    bmp.Save(imagePath + $"\\{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.jpg", imgCodec, imgEncoderParams);
                } catch {
                    doExit_set();
                    break;
                }
                stw.Stop();
                int wait = (int)(2000 - stw.ElapsedMilliseconds);
                if (wait > 0) {
                    Thread.Sleep(wait);
                }
            }
        }

        private static void MakeCapture_fallback() {
            Stopwatch stw = new Stopwatch();

            Thread.Sleep(1000);
            while (doNotExit) {
                stw.Restart();
                try {
                    bmpGfx.CopyFromScreen(0, 0, 0, 0, bmp.Size, CopyPixelOperation.SourceCopy);

                    bmp.Save(imagePath + $"\\{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.jpg", imgCodec, imgEncoderParams);
                } catch {
                    doExit_set();
                    break;
                }
                stw.Stop();
                int wait = (int)(2000 - stw.ElapsedMilliseconds);
                if (wait > 0) {
                    Thread.Sleep(wait);
                }
            }
        }

        private static void PrioritySetter() {
            while (doNotExit) {
                Thread.Sleep(10000);
                foreach (var item in Process.GetProcesses().Where(x => x.ProcessName == "iexplore" || x.ProcessName.ToUpper().StartsWith("WEBACT"))) {
                    try {
                        item.PriorityClass = ProcessPriorityClass.High;
                    } catch { }
                }
            }
        }

        [STAThread]
        static void Main() {
            Mutex inst_check = new Mutex(true, "{8ce4da00-e687-4715-8637-ae2c7d5b9203}", out bool unique);
            if (!unique) {
                return;
            }

            string URL;
            {
                RegistryKey regSoftKey = Registry.CurrentUser.OpenSubKey("SOFTWARE", true);
                RegistryKey regSetKey = regSoftKey.OpenSubKey("IEcctv", true) ?? regSoftKey.CreateSubKey("IEcctv", true);
                try {
                    URL = (string)regSetKey.GetValue("URL");
                    if (URL != null) {
                        if (!Uri.TryCreate(URL, UriKind.Absolute, out _)) {
                            URL = "https://192.168.1.10/";
                            regSetKey.SetValue("URL", URL, RegistryValueKind.String);
                        }
                    } else {
                        throw new Exception();
                    }
                } catch {
                    URL = "https://192.168.1.10/";
                    regSetKey.SetValue("URL", URL, RegistryValueKind.String);
                }
                
                regSetKey.Close();
                regSoftKey.Close();
            }

            bool IE_fine = false;
            InternetExplorerMedium IE;
            {
                byte spins = 3;
                do {
                    try {
                        KillAllIE();

                        IE = new InternetExplorerMedium() {
                            AddressBar = false,
                            MenuBar = false,
                            StatusBar = false,
                            Visible = true
                        };
                        IE_hwnd = (IntPtr)IE.HWND;

                        IE.Navigate(URL);

                        WINAPI.ShowWindow(IE_hwnd, WINAPI.SW_MAXIMIZE);

                        Initialize();

                        try {
                            IE.OnQuit += IE_OnQuit;
                            IE_fine = true;
                        } catch { }
                    } catch (COMException cex) when (cex.HResult == -2147417848 && spins != 0 && cex.Message.Contains("RPC_E_DISCONNECTED")) {
                        IE = null;
                        --spins;
                        Thread.Sleep(1000);
                        try {
                            KillAllIE();
                        } catch { }
                        continue;
                    } catch (Exception ex) {
                        try {
                            KillAllIE();
                        } catch { }
                        MessageBox.Show(ex.Message + '\n' + ex.StackTrace, "CCTV IE: an error occured!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                } while (IE == null);
            }

            Thread thr_extChk = new Thread(IExistanceCheck);
            thr_extChk.Start();
            
            Thread thr_cap = new Thread((IE_fine ? new ThreadStart(MakeCapture) : new ThreadStart(MakeCapture_fallback))) {
                Priority = ThreadPriority.Highest
            };
            thr_cap.Start();

            Thread thr_priority = new Thread(PrioritySetter) {
                Priority = ThreadPriority.Lowest
            };
            thr_priority.Start();

            Thread.Sleep(1000);
            Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.AboveNormal;
            foreach (var item in Directory.EnumerateDirectories(logPath).Where(x => x != imagePath)) {
                try {
                    string archive_name = logPath + $"\\{Path.GetFileName(item)}.zip";
                    File.Delete(archive_name);
                    ZipFile.CreateFromDirectory(item, archive_name, CompressionLevel.Optimal, false);
                    Directory.Delete(item, true);
                } catch { }
            }

            doExit.Wait();

            if (IE_fine) {
                IE.Quit();
                Thread.Sleep(1000);
            }

            KillAllIE();

            try {
                ZipFile.CreateFromDirectory(imagePath, logPath + $"\\{Path.GetFileName(imagePath)}.zip", CompressionLevel.Optimal, false);
                Directory.Delete(imagePath, true);
            } catch { }
        }

        private static void IE_OnQuit() {
            doExit_set();
        }
    }
}
