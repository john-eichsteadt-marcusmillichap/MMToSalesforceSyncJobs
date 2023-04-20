using System;

namespace SalesforceSyncJobs.Helpers
{
    public class LastRunTime
    {
        public static DateTime Get(DateTime currentDateTime, object lastRunTime)
        {
            DateTime runTime;
            if(!DateTime.TryParse(lastRunTime.ToString(), out runTime)) {
                runTime = currentDateTime;
            }

            return runTime;
        }
    }
}
