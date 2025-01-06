using System.Collections.Specialized;
using System.Management;
using System.Windows;

namespace PrintSpoolerWPF.PrintManagement.WMI
{
    public class PrinterMonitoring
    {
        List<PausedJob> pausedJobs = new List<PausedJob>();

        public void MonitorPrinterJobs()
        {
            Task.Run(async delegate
            {
                while (true)
                {
                    try
                    {
                        string searchQuery = "SELECT * FROM Win32_PrintJob";
                        ManagementObjectSearcher searchPrintJobs = new(searchQuery);
                        ManagementObjectCollection prntJobCollection = searchPrintJobs.Get();
                        foreach (ManagementObject prntJob in prntJobCollection)
                        {
                            string jobName = prntJob.Properties["Name"].Value.ToString();
                            //Job name would be of the format [Printer name], [Job ID]
                            char[] splitArr = new char[1];
                            splitArr[0] = Convert.ToChar(",");
                            string prnterName = jobName.Split(splitArr)[0];
                            int prntJobID = Convert.ToInt32(jobName.Split(splitArr)[1]);

                            string jobStatus = Convert.ToString(prntJob.Properties["JobStatus"]?.Value);

                            if (jobStatus == "Spooling")
                            {
                                PausePrintJob(prnterName, prntJobID);
                                pausedJobs.Add(new PausedJob()
                                {
                                    JobName = jobName,
                                    JobID = prntJobID
                                });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
                        break;
                    }
                    await Task.Delay(100);
                }
            });
        }

        public static bool PausePrintJob(string printerName, int printJobID)
        {
            bool isActionPerformed = false;
            try
            {
                string searchQuery = "SELECT * FROM Win32_PrintJob";
                ManagementObjectSearcher searchPrintJobs =
                         new ManagementObjectSearcher(searchQuery);
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
                            prntJob.InvokeMethod("Pause", null);
                            isActionPerformed = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return isActionPerformed;
        }

        public static bool ResumePrintJob(string printerName, int printJobID)
        {
            bool isActionPerformed = false;
            try
            {
                string searchQuery = "SELECT * FROM Win32_PrintJob";
                ManagementObjectSearcher searchPrintJobs =
                         new ManagementObjectSearcher(searchQuery);
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
                            prntJob.InvokeMethod("Resume", null);
                            isActionPerformed = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return isActionPerformed;
        }

        public static bool CancelPrintJob(string printerName, int printJobID)
        {
            bool isActionPerformed = false;
            try
            {
                string searchQuery = "SELECT * FROM Win32_PrintJob";
                ManagementObjectSearcher searchPrintJobs =
                       new ManagementObjectSearcher(searchQuery);
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
                            //performs a action similar to the cancel 
                            //operation of windows print console
                            prntJob.Delete();
                            isActionPerformed = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return isActionPerformed;
        }

        public static StringCollection GetPrintersCollection()
        {
            StringCollection printerNameCollection = new StringCollection();
            try
            {
                string searchQuery = "SELECT * FROM Win32_Printer";
                ManagementObjectSearcher searchPrinters =
                      new ManagementObjectSearcher(searchQuery);
                ManagementObjectCollection printerCollection = searchPrinters.Get();
                foreach (ManagementObject printer in printerCollection)
                {
                    printerNameCollection.Add(printer.Properties["Name"].Value.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return printerNameCollection;
        }

        public static StringCollection GetPrintJobsCollection(string printerName)
        {
            StringCollection printJobCollection = new StringCollection();
            try
            {
                string searchQuery = "SELECT * FROM Win32_PrintJob";

                /*searchQuery can also be mentioned with where Attribute,
                    but this is not working in Windows 2000 / ME / 98 machines 
                    and throws Invalid query error*/
                ManagementObjectSearcher searchPrintJobs =
                          new ManagementObjectSearcher(searchQuery);
                ManagementObjectCollection prntJobCollection = searchPrintJobs.Get();
                foreach (ManagementObject prntJob in prntJobCollection)
                {
                    System.String jobName = prntJob.Properties["Name"].Value.ToString();

                    //Job name would be of the format [Printer name], [Job ID]
                    char[] splitArr = new char[1];
                    splitArr[0] = Convert.ToChar(",");
                    string prnterName = jobName.Split(splitArr)[0];
                    string documentName = prntJob.Properties["Document"].Value.ToString();
                    if (String.Compare(prnterName, printerName, true) == 0)
                    {
                        printJobCollection.Add(documentName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
            return printJobCollection;
        }

    }
}
