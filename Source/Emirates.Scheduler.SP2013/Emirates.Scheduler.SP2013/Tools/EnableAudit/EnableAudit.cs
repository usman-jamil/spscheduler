using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Xml.Serialization;
using Microsoft.SharePoint;

namespace Emirates.Scheduler.SP2013.Tools
{
    public class EnableAudit : iTool
    {
        StringBuilder output = null;

        public EnableAudit()
        {
            output = new StringBuilder();
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);
            string inputXml = job.DownloadAttachment();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputXml);

            XmlNode rootNode = xmlDoc.SelectSingleNode("Sites");

            XmlNodeList siteNodes = xmlDoc.SelectNodes("//Site");
            foreach (XmlNode siteNode in siteNodes)
            {
                string url = siteNode.Attributes["Target"].Value;

                output.Append(string.Format("updating web: {0}" + Environment.NewLine, url));
                EnableAuditSettings(url);
            }

            string tmpFile = Scheduler.Instance.CreateTmpFile();

            System.IO.File.WriteAllText(tmpFile, output.ToString());

            result.AddFile(tmpFile);
            return result;
        }

        private void EnableAuditSettings(string site)
        {
            string auditLogLib = "ForSiteAuditing";
            string auditLogPath = "/" + auditLogLib + "/";

            try
            {
                using (SPWeb webNew = new SPSite(site).OpenWeb())
                {
                    SPSite spSite = webNew.Site;
                    Guid id = webNew.Lists.Add(auditLogLib, "For Auditing", SPListTemplateType.DocumentLibrary);

                    SPList listdoc = webNew.Lists[id];
                    listdoc.OnQuickLaunch = false;
                    listdoc.Hidden = true;
                    listdoc.Update();

                    SPDocumentLibrary _MyDocLibrary = (SPDocumentLibrary)webNew.Lists[auditLogLib];
                    SPFolderCollection _MyFolders = webNew.Folders;
                    string folderURL = spSite.Url + auditLogPath + "AuditReports";
                    _MyFolders.Add(folderURL);  //"ttp://adfsaccount:2222/My%20Documents/" + txtUpload.Text + "/");

                    _MyDocLibrary.Update();


                    spSite.TrimAuditLog = true;
                    spSite.AuditLogTrimmingRetention = 90;
                    Microsoft.Office.RecordsManagement.Reporting.AuditLogTrimmingReportCallout.SetAuditReportStorageLocation(spSite, folderURL);
                    //AuditLogTrimmingReportCallout.SetAuditReportStorageLocation

                    //   Audi
                    spSite.Audit.AuditFlags = SPAuditMaskType.Delete | SPAuditMaskType.Update | SPAuditMaskType.SecurityChange | SPAuditMaskType.SecurityChange | SPAuditMaskType.SchemaChange | SPAuditMaskType.Search;

                    //  site.Audit.AuditFlags =(Microsoft.SharePoint.SPAuditMaskType)auditOptions;
                    spSite.Audit.Update();
                }
            }
            catch { }
        }
    }
}
