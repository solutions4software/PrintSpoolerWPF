using System.Collections.ObjectModel;
using System.Timers;
using System.Windows;

namespace PrintSpoolerWPF.PrintManagement.Utilities
{
    public class WatchMonitoredPrinters
    {
        ObservableCollection<MonitorPrinter> monitorPrinters = new();

        readonly System.Timers.Timer timer;

        MainWindow window;

        public WatchMonitoredPrinters(MainWindow window)
        {
            try
            {
                if(timer == null)
                {
                    timer = new System.Timers.Timer();
                    timer.Elapsed += new ElapsedEventHandler(Timer_Tick);
                    timer.Interval = 1 * 1000;
                    timer.Enabled = true;
                }
                this.window = window;
            }
            catch (Exception ex)
            {
                //Logger.Error(ex);
            }
        }

        private void Timer_Tick(object? sender, ElapsedEventArgs e)
        {
            string printers = "Microsoft Print to PDF:Brother MFC-L5710DW series Printer";//RegistryManager.ReadMonitoredPrinters();
            
            if (string.IsNullOrEmpty(printers)) return;

            try
            {
                string[] printersList = printers.Split(':');
                foreach (var printer in printersList)
                {
                    if (!monitorPrinters.Any(p => p.printer == printer) && !string.IsNullOrEmpty(printer))
                    {
                        //MessageBox.Show($"{printer} added");
                        monitorPrinters.Add(new MonitorPrinter(printer, window));
                    }
                }
            }
            catch (Exception ex) 
            { 
                //Logger.Error(ex); 
            }
        }

    }
}
