using PrintSpoolerWPF.PrintManagement.Utilities;
using PrintSpoolerWPF.PrintManagement.WMI;
using System.Diagnostics;
using System.Management;
using System.Printing;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace PrintSpoolerWPF
{
    public partial class MainWindow : Window
    {
        private readonly DispatcherTimer _activityTimer;
        private Point _inactiveMousePosition = new Point(0, 0);

        PrinterMonitoring printerMonitoring = new();
        public MainWindow()
        {
            InitializeComponent();
            WatchMonitoredPrinters monitoredPrinters = new(this);
            //_ = new IdleTimeDetector(this);

            InputManager.Current.PreProcessInput += OnActivity;
            _activityTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10), IsEnabled = true };
            _activityTimer.Tick += OnInactivity;
        }

        void OnInactivity(object sender, EventArgs e)
        {
            // remember mouse position
            _inactiveMousePosition = Mouse.GetPosition(MainGrid);

            // set UI on inactivity
            _activityTimer.Stop();
            _activityTimer.Tick -= OnInactivity;
            MessageBox.Show("No activity detected for 10 seconds.");
            Window1 window1 = new();
            window1.Show();
            window1.Activate();
            Close();
        }

        void OnActivity(object sender, PreProcessInputEventArgs e)
        {
            InputEventArgs inputEventArgs = e.StagingItem.Input;

            if (inputEventArgs is MouseEventArgs || inputEventArgs is KeyboardEventArgs)
            {
                if (e.StagingItem.Input is MouseEventArgs)
                {
                    MouseEventArgs mouseEventArgs = (MouseEventArgs)e.StagingItem.Input;

                    // no button is pressed and the position is still the same as the application became inactive
                    if (mouseEventArgs.LeftButton == MouseButtonState.Released &&
                        mouseEventArgs.RightButton == MouseButtonState.Released &&
                        mouseEventArgs.MiddleButton == MouseButtonState.Released &&
                        mouseEventArgs.XButton1 == MouseButtonState.Released &&
                        mouseEventArgs.XButton2 == MouseButtonState.Released &&
                        _inactiveMousePosition == mouseEventArgs.GetPosition(MainGrid))
                        return;
                }

                // set UI on activity
                //rectangle.Visibility = Visibility.Visible;

                _activityTimer.Stop();
                _activityTimer.Start();
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
        }

        private void btnStartMonitoring_Click(object sender, RoutedEventArgs e)
        {
            printerMonitoring.MonitorPrinterJobs();
        }

        public void UpdatePrintInfoList(PrintJobChangeEventArgs e)
        {
            try
            {
                PrintSystemJobInfo job = e.JobInfo;
                Action invoker = () =>
                {
                    lbStatus.Items.Add(e.JobID + " - " + e.JobName + " - " + e.JobStatus);
                };
                if (job != null)
                {
                    lbStatus.Dispatcher.BeginInvoke(() => { invoker(); });
                }
                else
                {
                    //invoker();
                }
            }
            catch (Exception)
            {

            }
        }

        private void RefreshPrintQueue(PrintSystemJobInfo job)
        {
            PrintQueue pq = job.HostingPrintQueue;
            pq.Refresh();
        }

        #region DeletePrintJob

        /// <summary>
        /// Cancel the print job. This functions accepts the job number.
        /// An exception will be thrown if access denied.
        /// </summary>
        /// <param name="printJobID">int: Job number to cancel printing for.</param>
        /// <returns>bool: true if cancel successfull, else false.</returns>
        public bool DeletePrintJob(int printJobID)
        {
            // Variable declarations.
            bool isActionPerformed = false;
            string searchQuery;
            String jobName;
            char[] splitArr;
            int prntJobID;
            ManagementObjectSearcher searchPrintJobs;
            ManagementObjectCollection prntJobCollection;
            try
            {
                // Query to get all the queued printer jobs.
                searchQuery = "SELECT * FROM Win32_PrintJob";
                // Create an object using the above query.
                searchPrintJobs = new ManagementObjectSearcher(searchQuery);
                // Fire the query to get the collection of the printer jobs.
                prntJobCollection = searchPrintJobs.Get();

                // Look for the job you want to delete/cancel.
                foreach (ManagementObject prntJob in prntJobCollection)
                {
                    jobName = prntJob.Properties["Name"].Value.ToString();
                    // Job name would be of the format [Printer name], [Job ID]
                    splitArr = new char[1];
                    splitArr[0] = Convert.ToChar(",");
                    // Get the job ID.
                    prntJobID = Convert.ToInt32(jobName.Split(splitArr)[1]);
                    // If the Job Id equals the input job Id, then cancel the job.
                    if (prntJobID == printJobID)
                    {
                        // Performs a action similar to the cancel
                        // operation of windows print console
                        prntJob.Delete();
                        isActionPerformed = true;
                        break;
                    }
                }
                return isActionPerformed;
            }
            catch (Exception sysException)
            {
                // Log the exception.
                return false;
            }
        }

        #endregion DeletePrintJob

        public static bool JobDetails(string printerName, int printJobID)
        {
            bool isContinuePrinting = false;
            try
            {
                string searchQuery = "SELECT * FROM Win32_PrintJob";
                ManagementObjectSearcher searchPrintJobs = new ManagementObjectSearcher(searchQuery);
                ManagementObjectCollection prntJobCollection = searchPrintJobs.Get();

                foreach (ManagementObject prntJob in prntJobCollection)
                {
                    System.String jobName = prntJob.Properties["Name"].Value.ToString();

                    //Job name would be of the format [Printer name], [Job ID]
                    char[] splitArr = new char[1];
                    splitArr[0] = Convert.ToChar(",");
                    string prnterName = jobName.Split(splitArr)[0];
                    int prntJobID = Convert.ToInt32(jobName.Split(splitArr)[1]);
                    string documentName = prntJob.Properties["Document"].Value.ToString();

                    if (String.Compare(prnterName, printerName, true) == 0)
                    {
                        if (prntJobID == printJobID)
                        {
                            // MessageBox.Show("PAGINAS : " + prntJob.Properties["TotalPages"].Value.ToString() + documentName + " " + prntJobID);
                            //prntJob.InvokeMethod("Pause", null);

                            string jobStatus = prntJob.Properties["JobStatus"].Value + "";
                            int totalPages = (int)((uint)prntJob.Properties["TotalPages"].Value);
                            string color = (prntJob.Properties["Color"].Value + "" == "Color") ? "Color" : "Black and white";

                            string jobDetails = string.Format(
                                    "Job Status: {0}\n" +
                                    "Printer: {1}\n" +
                                    "No. of Pages: {2}\n" +
                                    "Color: {3}\n" +
                                    "\nWould you like to continue printing?",
                                    jobStatus,
                                    prnterName,
                                    totalPages,
                                    color);
                            //Logger.Info(jobDetails);
                            MessageBoxResult result = MessageBox.Show(jobDetails, "Job Details", MessageBoxButton.YesNo);
                            if (result == MessageBoxResult.Yes)
                            {
                                isContinuePrinting = true;
                            }

                            //prntJob.InvokeMethod("Resume", null);

                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //Logger.Error(ex);
                //MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }

            return isContinuePrinting;
        }

    }
}