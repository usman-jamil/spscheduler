using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Xml.Serialization;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration.Claims;

namespace Emirates.Scheduler.SP2013.Tools
{
    public class DeleteAlerts : iTool
    {
        StringBuilder output = null;

        public DeleteAlerts()
        {
            output = new StringBuilder();
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);
            string inputXml = job.DownloadAttachment();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputXml);

            XmlNodeList siteNodes = xmlDoc.SelectNodes("notifications/site");
            foreach (XmlNode siteNode in siteNodes)
            {
                string sourceUrl = siteNode.Attributes["source"].Value;
                string url = siteNode.Attributes["target"].Value;

                output.Append(string.Format("updating web: {0,20}" + Environment.NewLine, url));
                RemoveAlerts(siteNode, url);
            }

            string tmpFile = Scheduler.Instance.CreateTmpFile();

            System.IO.File.WriteAllText(tmpFile, output.ToString());

            result.AddFile(tmpFile);
            return result;
        }

        private void RemoveAlerts(XmlNode siteNode, string url)
        {
            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                SPClaimProviderManager cpm = SPClaimProviderManager.Local;
                SPClaim userClaim = cpm.ConvertIdentifierToClaim("emirates\\s717981", SPIdentifierTypes.WindowsSamAccountName);
                SPUser tempUser = web.EnsureUser(userClaim.ToEncodedString());

                web.AllowUnsafeUpdates = true;
                try
                {
                    SPAlertCollection allAlerts = web.Alerts;

                    List<Guid> alertsToDelete = new List<Guid>();

                    foreach (SPAlert spAlert in allAlerts)
                    {
                        alertsToDelete.Add(spAlert.ID);
                    }

                    Guid []alerts = alertsToDelete.ToArray();
                    for (int i = 0; i < alerts.Length; i++)
                    {
                        SPAlert alert = allAlerts[alerts[i]];
                        alert.User = tempUser;
                        alert.Status = SPAlertStatus.Off;
                        alert.Update();
                    }

                    foreach (Guid alertGuid in alertsToDelete)
                    {
                        allAlerts.Delete(alertGuid);
                    }

                    web.Update();
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
                web.AllowUnsafeUpdates = false;
            }
        }
    }
}
