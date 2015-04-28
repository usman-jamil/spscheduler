using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Emirates.Scheduler.SP2007.Tools;

namespace Emirates.Scheduler.SP2007
{
    public class Job
    {
        private int id;
        private int prerequisite;
        private string title;
        private string operation;
        private DateTime when;
        private string[] attachments;
        private iTool associatedTool;

        public Job(int id, int prerequisite, string title, string operation, DateTime when, SPAttachmentCollection attachments)
        {
            this.id = id;
            this.prerequisite = prerequisite;
            this.title = title;
            this.operation = operation;
            this.when = when;
            this.attachments = AddAttachments(attachments);
            this.associatedTool = IdentifyTool();
        }

        public Job(Job job)
        {
            this.id = job.Id;
            this.prerequisite = job.Prerequisite;
            this.title = job.Title;
            this.operation = job.operation;
            this.when = job.when;
            this.associatedTool = IdentifyTool();
        }

        private iTool IdentifyTool()
        {
            iTool toolToExecute = new NOP();

            switch (this.operation.ToLower())
            {
                case "export security":
                    toolToExecute = new ExportSecurity();
                    break;

                case "export alerts":
                    toolToExecute = new ExportAlerts();
                    break;

                case "comparison report":
                    toolToExecute = new ComparisonReport();
                    break;

                case "site structure":
                    toolToExecute = new SiteStructure();
                    break;

                case "list report":
                    toolToExecute = new ListReport();
                    break;
            }

            return toolToExecute;
        }

        public string[] AddAttachments(SPAttachmentCollection spAttachments)
        {
            List<string> files = new List<string>();

            for (int i = 0; i < spAttachments.Count; i++)
            {
                files.Add(spAttachments[i]);
            }

            return files.ToArray();
        }

        public string DownloadAttachment()
        {
            string sWeb = System.Configuration.ConfigurationManager.AppSettings["Server"];
            string sList = System.Configuration.ConfigurationManager.AppSettings["List"];
            string content = string.Empty;

            using (SPWeb web = new SPSite(sWeb).OpenWeb())
            {
                SPList list = web.Lists[sList];
                SPListItem item = list.GetItemById(this.id);
                return DownloadAttachment(item.Attachments[0]);
            }
        }

        public string DownloadAttachment(string attachment)
        {
            string sWeb = System.Configuration.ConfigurationManager.AppSettings["Server"];
            string sList = System.Configuration.ConfigurationManager.AppSettings["List"];
            string tempXml = string.Empty;

            using (SPWeb web = new SPSite(sWeb).OpenWeb())
            {
                SPList list = web.Lists[sList];

                string fileUrl = string.Format("{0}/Attachments/{1}/{2}",
                    list.RootFolder.ServerRelativeUrl,
                    this.id,
                    attachment);
                SPFile file = web.GetFile(fileUrl);
                byte[] data = file.OpenBinary();
                
                //content = System.Text.Encoding.UTF8.GetString(data);
                tempXml = Scheduler.Instance.CreateTmpFile();
                System.IO.File.WriteAllBytes(tempXml, data);
            }

            return tempXml;
        }

        public int Id
        {
            get
            {
                return this.id;
            }
        }

        public int Prerequisite
        {
            get
            {
                return this.prerequisite;
            }
        }

        public string Title
        {
            get
            {
                return this.title;
            }
        }

        public string Operation
        {
            get
            {
                return this.operation;
            }
        }

        public DateTime When
        {
            get
            {
                return this.when;
            }
        }

        public string[] Attachments
        {
            get
            {
                return this.attachments;
            }
        }

        public iTool AssociatedTool
        {
            get
            {
                return this.associatedTool;
            }
        }
    }
}
