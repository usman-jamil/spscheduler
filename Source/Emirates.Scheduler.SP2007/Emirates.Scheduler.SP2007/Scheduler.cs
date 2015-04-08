using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.SharePoint;

namespace Emirates.Scheduler.SP2007
{
    public sealed class Scheduler
    {
        static readonly Scheduler instance = new Scheduler();

        /* ======================== */

        // Explicit static constructor to tell C# compiler
        // not to mark type as beforefieldinit
        static Scheduler()
        {
        }

        Scheduler()
        {
        }

        public static Scheduler Instance
        {
            get
            {
                return instance;
            }
        }

        public Job[] DownloadNewJobs()
        {
            List<Job> jobs = new List<Job>();
            string sWeb = System.Configuration.ConfigurationManager.AppSettings["Server"];
            string sList = System.Configuration.ConfigurationManager.AppSettings["List"];
            bool moderationEnabled = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["ModerationEnabled"]);
            uint itemsToProcess = Convert.ToUInt32(System.Configuration.ConfigurationManager.AppSettings["ItemsToProcess"]);

            using (SPWeb web = new SPSite(sWeb).OpenWeb())
            {
                SPList list = web.Lists[sList];

                string camlQuery = string.Empty;

                if (!moderationEnabled)
                {
                    camlQuery = @"
                    <OrderBy>
                        <FieldRef Name=""ID"" />
                    </OrderBy>
                    <Where>
                        <And>
                            <Eq>
                                <FieldRef Name=""Status"" />
                                <Value Type=""Text"">Not Started</Value>
                            </Eq>
                            <Lt>
                                <FieldRef Name=""When"" />
                                <Value IncludeTimeValue='TRUE' Type=""DateTime"">
                                    <Today />
                                </Value>
                            </Lt>
                        </And>
                    </Where>";
                }
                else
                {
                    camlQuery = @"
                    <OrderBy>
                        <FieldRef Name=""ID"" />
                    </OrderBy>
                    <Where>
                        <And>
                            <And>
                                <Eq>
                                    <FieldRef Name=""Status"" />
                                    <Value Type=""Text"">Not Started</Value>
                                </Eq>
                                <Lt>
                                    <FieldRef Name=""When"" />
                                    <Value IncludeTimeValue='TRUE' Type=""DateTime""><Today /></Value>
                                </Lt>
                            </And>
                                <Eq>
                                    <FieldRef Name=""_ModerationStatus"" />
                                    <Value Type=""ModStat"">Approved</Value>
                                </Eq>
                            </And>
                    </Where>";
                }
                SPQuery query = new SPQuery();
                query.Query = camlQuery;
                query.RowLimit = itemsToProcess;

                SPListItemCollection results = list.GetItems(query);
                foreach (SPListItem item in results)
                {
                    string pre = Convert.ToString(item["Prerequisite"]);

                    Job job = new Job(item.ID,
                        string.IsNullOrEmpty(pre) ? 0 : Convert.ToInt32(pre),
                        item.Title,
                        Convert.ToString(item["Task"]),
                        Convert.ToDateTime(item["When"]),
                        item.Attachments);

                    jobs.Add(job);
                }
            }

            return jobs.ToArray();
        }

        public void CompleteJob(Result result)
        {
            string sWeb = System.Configuration.ConfigurationManager.AppSettings["Server"];
            string sList = System.Configuration.ConfigurationManager.AppSettings["List"];
            uint itemsToProcess = Convert.ToUInt32(System.Configuration.ConfigurationManager.AppSettings["ItemsToProcess"]);
            string[] outputFiles = result.GetAllFiles();

            using (SPWeb web = new SPSite(sWeb).OpenWeb())
            {
                SPList list = web.Lists[sList];

                web.AllowUnsafeUpdates = true;

                SPListItem item = list.GetItemById(result.Id);
                item["Status"] = "Completed";
                item["Last Completed"] = DateTime.Now;
                item.Update();

                foreach (string outputFile in outputFiles)
                {
                    SPListItem updatedItem = list.GetItemById(result.Id);
                    byte[] data = File.ReadAllBytes(outputFile);
                    FileInfo fileInfo = new FileInfo(outputFile);
                    updatedItem.Attachments.Add(fileInfo.Name, data);
                    updatedItem.Update();
                }

                web.AllowUnsafeUpdates = false;
            }
        }

        public void InitiateJob(Job job)
        {
            string sWeb = System.Configuration.ConfigurationManager.AppSettings["Server"];
            string sList = System.Configuration.ConfigurationManager.AppSettings["List"];
            uint itemsToProcess = Convert.ToUInt32(System.Configuration.ConfigurationManager.AppSettings["ItemsToProcess"]);

            using (SPWeb web = new SPSite(sWeb).OpenWeb())
            {
                SPList list = web.Lists[sList];

                web.AllowUnsafeUpdates = true;

                SPListItem item = list.GetItemById(job.Id);
                item["Status"] = "In Progress";
                item.Update();

                web.AllowUnsafeUpdates = false;
            }
        }

        public bool IsJobReady(Job job)
        {
            //Check if all dependencies are resolved
            bool ready = (job.Prerequisite == 0);
            if (job.Prerequisite > 0)
            {
                string sWeb = System.Configuration.ConfigurationManager.AppSettings["Server"];
                string sList = System.Configuration.ConfigurationManager.AppSettings["List"];
                uint itemsToProcess = Convert.ToUInt32(System.Configuration.ConfigurationManager.AppSettings["ItemsToProcess"]);

                using (SPWeb web = new SPSite(sWeb).OpenWeb())
                {
                    SPList list = web.Lists[sList];

                    try
                    {
                        SPListItem item = list.GetItemById(job.Prerequisite);
                        string jobStatus = Convert.ToString(item["Status"]);
                        ready = jobStatus.ToLower().Equals("completed");
                    }
                    catch { }
                }
            }

            return ready;
        }

        public string CreateTmpFile()
        {
            string fileName = string.Empty;

            try
            {
                // Get the full name of the newly created Temporary file. 
                // Note that the GetTempFileName() method actually creates
                // a 0-byte file and returns the name of the created file.
                fileName = Path.GetTempFileName();

                // Craete a FileInfo object to set the file's attributes
                FileInfo fileInfo = new FileInfo(fileName);

                // Set the Attribute property of this file to Temporary. 
                // Although this is not completely necessary, the .NET Framework is able 
                // to optimize the use of Temporary files by keeping them cached in memory.
                fileInfo.Attributes = FileAttributes.Temporary;
            }
            catch { }

            return fileName;
        }
    }
}
