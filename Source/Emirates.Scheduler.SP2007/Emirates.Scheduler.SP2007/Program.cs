using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;

namespace Emirates.Scheduler.SP2007
{
    class Program
    {
        static void Main(string[] args)
        {
            bool createdNew;

            string mutexId = System.Configuration.ConfigurationManager.AppSettings["Mutex"];
            Mutex m = new Mutex(true, mutexId, out createdNew);

            if (createdNew)
            {
                Scheduler scheduler = Scheduler.Instance;
                Job []jobs = scheduler.DownloadNewJobs();

                foreach (Job job in jobs)
                {
                    //Check for dependencies
                    if (scheduler.IsJobReady(job))
                    {
                        scheduler.InitiateJob(job);
                        Result result = job.AssociatedTool.Execute(job);
                        scheduler.CompleteJob(result);
                    }
                }
            }
        }
    }
}
