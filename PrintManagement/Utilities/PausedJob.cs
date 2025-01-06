using System.Printing;

namespace PrintSpoolerWPF.PrintManagement.Utilities
{
    internal class PausedJob
    {
        public int JobID { get; private set; }
        public long PausedTime { get; private set; }
        
        public PausedJob(int jobID)
        {
            JobID = jobID;
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
