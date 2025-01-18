using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace PrintSpoolerWPF
{
    public class IdleTimeDetector
    {
        Window parentWindow;

        public IdleTimeDetector(Window parentWindow)
        {
            this.parentWindow = parentWindow;
            DispatcherTimer timer = new();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += timer_Tick; 
            timer.Start();
        }

        private void timer_Tick(object? sender, EventArgs e)
        {
            try
            {
                var idleTime = GetIdleTimeInfo();

                if (idleTime == null) return;

                if (idleTime.IdleTime.TotalMilliseconds >= 10 * 1000)
                {
                    parentWindow.Close();
                }
            }
            catch (Exception)
            {

            }
        }

        [DllImport("user32.dll")]
        static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        public IdleTimeInfo GetIdleTimeInfo()
        {
            int systemUptime = Environment.TickCount, lastInputTicks = 0, idleTicks = 0;

            IdleTimeInfo idleTimeInfo = new();

            try
            {
                LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
                lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
                lastInputInfo.dwTime = 0;

                if (GetLastInputInfo(ref lastInputInfo))
                {
                    lastInputTicks = (int)lastInputInfo.dwTime;
                    idleTicks = systemUptime - lastInputTicks;
                }

                idleTimeInfo.LastInputTime = DateTime.Now.AddMilliseconds(-1 * idleTicks);
                idleTimeInfo.IdleTime = new TimeSpan(0, 0, 0, 0, idleTicks);
                idleTimeInfo.SystemUptimeMilliseconds = systemUptime;
            }
            catch (Exception)
            {
                
            }

            return idleTimeInfo;
        }
    }

    public class IdleTimeInfo
    {
        public DateTime LastInputTime { get; internal set; }
        public TimeSpan IdleTime { get; internal set; }
        public int SystemUptimeMilliseconds { get; internal set; }
    }

    internal struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }
}
