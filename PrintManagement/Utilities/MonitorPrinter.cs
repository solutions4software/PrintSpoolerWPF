using System.Management;
using System.Printing;
using System.Windows;

namespace PrintSpoolerWPF.PrintManagement.Utilities
{
    internal class MonitorPrinter
    {
        List<PausedJob> pausedJobs = new List<PausedJob>();
        PrintQueueMonitor pqm;

        readonly public string printer;

        MainWindow window;

        public MonitorPrinter(string printer, MainWindow window)
        {
            try
            {
                this.printer = printer;
                this.window = window;
                pqm = new PrintQueueMonitor(printer);
                pqm.OnJobStatusChange += new PrintJobStatusChanged(pqm_OnJobStatusChange);
            }
            catch (Exception ex)
            {
                //Logger.Error(ex);
            }
        }

        void pqm_OnJobStatusChange(object Sender, PrintJobChangeEventArgs e)
        {
            try
            {
                PrintSystemJobInfo job = e.JobInfo;

                JOBSTATUS JobStatus = e.JobStatus;

                if (job == null) return;

                window.UpdatePrintInfoList(job, e);

                if (job.IsSpooling)
                {
                    job.Pause();
                    //job.HostingPrintQueue.Pause();
                }
                else if (job.IsPaused)
                {
                    //job.HostingPrintQueue.Resume();

                    RefreshPrintQueue(job);

                    if (JobDetails(printer, e.JobID))
                    {
                        job.Resume();
                    }
                    else
                    {
                        DeletePrintJob(e.JobID);
                    }
                }
                else if(job.IsCompleted)
                {
                    DeletePrintJob(e.JobID);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                            string color = (prntJob.Properties["Color"].Value+"" == "Color") ? "Color" : "Black and white";

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
