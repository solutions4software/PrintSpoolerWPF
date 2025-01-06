using System.Printing;

namespace PrintSpoolerWPF.PrintManagement.WMI
{
    public class PausedJob
    {
        public string PrinterName { get; set; }
        public string JobName { get; set; }
        public int JobID { get; set; }
        public long PausedTime { get; private set; }
        
        public PausedJob()
        {
            PausedTime = 20 * 1000;
        }
        public PausedJob(int jobID, PrintSystemJobInfo jobInfo, long pausedTime)
        {
            JobID = jobID;
            PausedTime = pausedTime;
        }

        public void UpdateTime()
        {
            if (PausedTime > 0)
            {
                PausedTime -= 1000;
            } 
        }
    }
}
