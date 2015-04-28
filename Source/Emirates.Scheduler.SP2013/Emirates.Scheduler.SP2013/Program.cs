using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;

namespace Emirates.Scheduler.SP2013
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
                Job[] jobs = scheduler.DownloadNewJobs();

                foreach (Job job in jobs)
                {
                    Result result = null;
                    //Check for dependencies
                    if (scheduler.IsJobReady(job))
                    {
                        try
                        {
                            scheduler.InitiateJob(job);
                            result = job.AssociatedTool.Execute(job);
                        }
                        finally
                        {
                            scheduler.CompleteJob(result);
                        }
                    }
                }
            }
        }
    }
}
