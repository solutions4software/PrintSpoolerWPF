﻿using System.Printing;
using System.Runtime.InteropServices;
using System.Windows;

namespace PrintSpoolerWPF.PrintManagement.Utilities
{
    public class PrintJobChangeEventArgs : EventArgs
    {
        #region private variables
        private int _jobID = 0;
        private string _jobName = "";
        private JOBSTATUS _jobStatus = new JOBSTATUS();
        private PrintSystemJobInfo _jobInfo = null;
        #endregion

        public int JobID { get { return _jobID; } }
        public string JobName { get { return _jobName; } }
        public JOBSTATUS JobStatus { get { return _jobStatus; } }
        public PrintSystemJobInfo JobInfo { get { return _jobInfo; } }
        public PrintJobChangeEventArgs(int intJobID, string strJobName, JOBSTATUS jStatus, PrintSystemJobInfo objJobInfo)
            : base()
        {
            _jobID = intJobID;
            _jobName = strJobName;
            _jobStatus = jStatus;
            _jobInfo = objJobInfo;
        }
    }

    public delegate void PrintJobStatusChanged(object Sender, PrintJobChangeEventArgs e);

    public class PrintQueueMonitor
    {
        #region DLL Import Functions
        [DllImport("winspool.drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool OpenPrinter(string pPrinterName,
        out IntPtr phPrinter,
        int pDefault);

        [DllImport("winspool.drv", EntryPoint = "ClosePrinter",
            SetLastError = true,
            ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClosePrinter
        (int hPrinter);

        [DllImport("winspool.drv",
        EntryPoint = "FindFirstPrinterChangeNotification",
        SetLastError = true, CharSet = CharSet.Ansi,
        ExactSpelling = true,
        CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr FindFirstPrinterChangeNotification
                            ([In()] IntPtr hPrinter,
                            [In()] int fwFlags,
                            [In()] int fwOptions,
                            [In(), MarshalAs(UnmanagedType.LPStruct)] PRINTER_NOTIFY_OPTIONS pPrinterNotifyOptions);

        [DllImport("winspool.drv", EntryPoint = "FindNextPrinterChangeNotification",
        SetLastError = true, CharSet = CharSet.Ansi,
        ExactSpelling = false,
        CallingConvention = CallingConvention.StdCall)]
        public static extern bool FindNextPrinterChangeNotification
                            ([In()] IntPtr hChangeObject,
                             [Out()] out int pdwChange,
                             [In(), MarshalAs(UnmanagedType.LPStruct)] PRINTER_NOTIFY_OPTIONS pPrinterNotifyOptions,
                            [Out()] out IntPtr lppPrinterNotifyInfo
                                 );
        #endregion

        #region Constants
        const int PRINTER_NOTIFY_OPTIONS_REFRESH = 1;
        #endregion
        #region Events
        public event PrintJobStatusChanged OnJobStatusChange;
        #endregion

        #region private variables
        private IntPtr _printerHandle = IntPtr.Zero;
        private string _spoolerName = "";
        private ManualResetEvent _mrEvent = new ManualResetEvent(false);
        private RegisteredWaitHandle _waitHandle = null;
        private IntPtr _changeHandle = IntPtr.Zero;
        private PRINTER_NOTIFY_OPTIONS _notifyOptions = new PRINTER_NOTIFY_OPTIONS();
        private Dictionary<int, string> objJobDict = new Dictionary<int, string>();
        private PrintQueue _spooler = null;
        #endregion

        #region constructor
        public PrintQueueMonitor(string strSpoolName)
        {
            // Let us open the printer and get the printer handle.
            _spoolerName = strSpoolName;
            //Start Monitoring
            Start(strSpoolName);
        }
        #endregion

        #region destructor
        ~PrintQueueMonitor()
        {
            Stop();
        }
        #endregion

        #region StartMonitoring
        public void Start(string printer)
        {
            if (_spoolerName != printer) return;
            try
            {
                OpenPrinter(_spoolerName, out _printerHandle, 0);
                if (_printerHandle != IntPtr.Zero)
                {
                    //We got a valid Printer handle.  Let us register for change notification....
                    _changeHandle = FindFirstPrinterChangeNotification(_printerHandle, (int)PRINTER_CHANGES.PRINTER_CHANGE_JOB, 0, _notifyOptions);
                    // We have successfully registered for change notification.  Let us capture the handle...
                    _mrEvent.Handle = _changeHandle;
                    //Now, let us wait for change notification from the printer queue....
                    _waitHandle = ThreadPool.RegisterWaitForSingleObject(_mrEvent, new WaitOrTimerCallback(PrinterNotifyWaitCallback), _mrEvent, -1, true);
                }

                _spooler = new PrintQueue(new PrintServer(), _spoolerName);
                foreach (PrintSystemJobInfo psi in _spooler.GetPrintJobInfoCollection())
                {
                    objJobDict[psi.JobIdentifier] = psi.Name;
                }
            }
            catch (Exception ex)
            {
                //Logger.Error(ex);
            }
        }
        #endregion

        #region StopMonitoring
        public void Stop()
        {
            try
            {
                if (_printerHandle != IntPtr.Zero)
                {
                    ClosePrinter((int)_printerHandle);
                    _printerHandle = IntPtr.Zero;
                }
            }
            catch (Exception ex)
            {
                //Logger.Error(ex);
            }
        }
        #endregion


        #region Callback Function
        public void PrinterNotifyWaitCallback(object state, bool timedOut)
        {
            if (_printerHandle == IntPtr.Zero) return;
            #region read notification details
            _notifyOptions.Count = 1;
            int pdwChange = 0;
            IntPtr pNotifyInfo = IntPtr.Zero;
            try
            {
                bool bResult = FindNextPrinterChangeNotification(_changeHandle, out pdwChange, _notifyOptions, out pNotifyInfo);
                //If the Printer Change Notification Call did not give data, exit code
                if (bResult == false || pNotifyInfo == IntPtr.Zero) return;
                //If the Change Notification was not relgated to job, exit code
                bool bJobRelatedChange = (pdwChange & PRINTER_CHANGES.PRINTER_CHANGE_ADD_JOB) == PRINTER_CHANGES.PRINTER_CHANGE_ADD_JOB ||
                                         (pdwChange & PRINTER_CHANGES.PRINTER_CHANGE_SET_JOB) == PRINTER_CHANGES.PRINTER_CHANGE_SET_JOB ||
                                         (pdwChange & PRINTER_CHANGES.PRINTER_CHANGE_DELETE_JOB) == PRINTER_CHANGES.PRINTER_CHANGE_DELETE_JOB ||
                                         (pdwChange & PRINTER_CHANGES.PRINTER_CHANGE_WRITE_JOB) == PRINTER_CHANGES.PRINTER_CHANGE_WRITE_JOB;
                if (!bJobRelatedChange) return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            #endregion 

            #region populate Notification Information
            try
            {
                //Now, let us initialize and populate the Notify Info data
                PRINTER_NOTIFY_INFO info = (PRINTER_NOTIFY_INFO)Marshal.PtrToStructure(pNotifyInfo, typeof(PRINTER_NOTIFY_INFO));
                long pData = (long)pNotifyInfo + (long)Marshal.OffsetOf(typeof(PRINTER_NOTIFY_INFO), "aData");
                PRINTER_NOTIFY_INFO_DATA[] data = new PRINTER_NOTIFY_INFO_DATA[info.Count];
                for (uint i = 0; i < info.Count; i++)
                {
                    data[i] = (PRINTER_NOTIFY_INFO_DATA)Marshal.PtrToStructure((IntPtr)pData, typeof(PRINTER_NOTIFY_INFO_DATA));
                    pData += Marshal.SizeOf(typeof(PRINTER_NOTIFY_INFO_DATA));
                }

                #region iterate through all elements in the data array
                for (int i = 0; i < data.Count(); i++)
                {

                    if (data[i].Field == (ushort)PRINTERJOBNOTIFICATIONTYPES.JOB_NOTIFY_FIELD_STATUS &&
                         data[i].Type == (ushort)PRINTERNOTIFICATIONTYPES.JOB_NOTIFY_TYPE)
                    {
                        JOBSTATUS jStatus = (JOBSTATUS)Enum.Parse(typeof(JOBSTATUS), data[i].NotifyData.Data.cbBuf.ToString());
                        int intJobID = (int)data[i].Id;
                        string strJobName = "";
                        PrintSystemJobInfo pji = null;
                        try
                        {
                            _spooler = new PrintQueue(new PrintServer(), _spoolerName);
                            pji = _spooler.GetJob(intJobID);
                            if (!objJobDict.ContainsKey(intJobID))
                                objJobDict[intJobID] = pji.Name;
                            strJobName = pji.Name;
                        }
                        catch
                        {
                            pji = null;
                            objJobDict.TryGetValue(intJobID, out strJobName);
                            if (strJobName == null) strJobName = "";
                        }

                        if (OnJobStatusChange != null)
                        {
                            //Let us raise the event
                            OnJobStatusChange(this, new PrintJobChangeEventArgs(intJobID, strJobName, jStatus, pji));
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            #endregion

            #region reset the Event and wait for the next event
            try
            {
                _mrEvent.Reset();
                _waitHandle = ThreadPool.RegisterWaitForSingleObject(_mrEvent, new WaitOrTimerCallback(PrinterNotifyWaitCallback), _mrEvent, -1, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            #endregion 

        }
        #endregion

    }
}
