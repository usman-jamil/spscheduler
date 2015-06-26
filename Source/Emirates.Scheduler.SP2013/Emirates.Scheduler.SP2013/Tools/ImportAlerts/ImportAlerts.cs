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
    public class ImportAlerts : iTool
    {
        StringBuilder output = null;

        public ImportAlerts()
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
                output.Append(string.Format("updating lists" + Environment.NewLine));
                CheckLists(siteNode, url);
                output.Append(string.Format("updating folders" + Environment.NewLine));
                CheckFolders(siteNode, url);
                output.Append(string.Format("updating items" + Environment.NewLine));
                CheckItems(siteNode, url);
            }

            string tmpFile = Scheduler.Instance.CreateTmpFile();

            System.IO.File.WriteAllText(tmpFile, output.ToString());

            result.AddFile(tmpFile);
            return result;
        }

        private void CheckLists(XmlNode siteNode, string url)
        {
            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                XmlNodeList alertNodes = siteNode.SelectNodes("alert[@type='List']");
                foreach (XmlNode alertNode in alertNodes)
                {
                    string alertTitle = alertNode.Attributes["title"].Value;
                    string listTitle = alertNode.Attributes["list"].Value;
                    string loginName = alertNode.Attributes["user"].Value;
                    string objectType = alertNode.Attributes["object"].Value;
                    bool isFolder = objectType.ToLower().Equals("folder");

                    if (!isFolder)
                    {
                        output.Append(string.Format("user: {0,20}...", loginName));
                        try
                        {
                            SPList list = web.Lists[listTitle];
                            SPClaimProviderManager cpm = SPClaimProviderManager.Local;
                            SPClaim userClaim = cpm.ConvertIdentifierToClaim(loginName, SPIdentifierTypes.WindowsSamAccountName);
                            SPUser user = web.EnsureUser(userClaim.ToEncodedString());

                            string eventType = alertNode.Attributes["event"].Value;
                            SPEventType spEventType = (SPEventType)Enum.Parse(typeof(SPEventType), eventType);

                            string eventFrequency = alertNode.Attributes["frequency"].Value;
                            SPAlertFrequency spAlertFrequency = (SPAlertFrequency)Enum.Parse(typeof(SPAlertFrequency), eventFrequency);

                            string type = alertNode.Attributes["type"].Value;
                            SPAlertType spAlertType = (SPAlertType)Enum.Parse(typeof(SPAlertType), type);

                            SPAlert newAlert = user.Alerts.Add();
                            newAlert.Title = alertTitle;
                            newAlert.AlertType = spAlertType;
                            newAlert.List = list;
                            newAlert.DeliveryChannels = SPAlertDeliveryChannels.Email;
                            newAlert.EventType = spEventType;
                            newAlert.AlertFrequency = spAlertFrequency;
                            newAlert.Status = SPAlertStatus.On;
                            newAlert.Update(false);
                            output.Append(string.Format("Complete" + Environment.NewLine));
                        }
                        catch (Exception ex) { output.Append(string.Format("error: {0,20}" + Environment.NewLine, ex.Message)); }
                    }
                }
            }
        }

        private void CheckFolders(XmlNode siteNode, string url)
        {
            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                XmlNodeList alertNodes = siteNode.SelectNodes("alert[@type='List']");
                foreach (XmlNode alertNode in alertNodes)
                {
                    string alertTitle = alertNode.Attributes["title"].Value;

                    string listTitle = alertNode.Attributes["list"].Value;
                    string loginName = alertNode.Attributes["user"].Value; 
                    string objectType = alertNode.Attributes["object"].Value;
                    bool isFolder = objectType.ToLower().Equals("folder");

                    if (isFolder)
                    {
                        string itemUrl = alertNode.Attributes["url"].Value;
                        output.Append(string.Format("user: {0,20}...", loginName));
                        try
                        {
                            SPList list = web.Lists[listTitle];
                            SPClaimProviderManager cpm = SPClaimProviderManager.Local;
                            SPClaim userClaim = cpm.ConvertIdentifierToClaim(loginName, SPIdentifierTypes.WindowsSamAccountName);
                            SPUser user = web.EnsureUser(userClaim.ToEncodedString());

                            string eventType = alertNode.Attributes["event"].Value;
                            SPEventType spEventType = (SPEventType)Enum.Parse(typeof(SPEventType), eventType);

                            string eventFrequency = alertNode.Attributes["frequency"].Value;
                            SPAlertFrequency spAlertFrequency = (SPAlertFrequency)Enum.Parse(typeof(SPAlertFrequency), eventFrequency);

                            string type = alertNode.Attributes["type"].Value;
                            SPAlertType spAlertType = (SPAlertType)Enum.Parse(typeof(SPAlertType), type);

                            SPFolder folder = web.GetFolder(itemUrl);
                            SPListItem item = folder.Item;

                            SPAlert newAlert = user.Alerts.Add();

                            newAlert.Title = alertTitle;
                            newAlert.AlertType = SPAlertType.Item;
                            newAlert.Item = item;
                            newAlert.DeliveryChannels = SPAlertDeliveryChannels.Email;
                            newAlert.EventType = spEventType;
                            newAlert.AlertFrequency = spAlertFrequency;
                            newAlert.Status = SPAlertStatus.On;
                            newAlert.Update(false);
                            output.Append(string.Format("Complete" + Environment.NewLine));
                        }
                        catch (Exception ex) { output.Append(string.Format("error: {0,20}" + Environment.NewLine, ex.Message)); }
                    }
                }
            }
        }

        private void CheckItems(XmlNode siteNode, string url)
        {
            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                XmlNodeList alertNodes = siteNode.SelectNodes("alert[@type='Item']");
                foreach (XmlNode alertNode in alertNodes)
                {
                    string alertTitle = alertNode.Attributes["title"].Value;
                    string listTitle = alertNode.Attributes["list"].Value;
                    string loginName = alertNode.Attributes["user"].Value;
                    int itemId = Int32.Parse(alertNode.Attributes["id"].Value);
                    string itemUrl = alertNode.Attributes["url"].Value;
                    string objectType = alertNode.Attributes["object"].Value;

                    output.Append(string.Format("user: {0,20}...", loginName));
                    try
                    {
                        SPList list = web.Lists[listTitle];
                        SPClaimProviderManager cpm = SPClaimProviderManager.Local;
                        SPClaim userClaim = cpm.ConvertIdentifierToClaim(loginName, SPIdentifierTypes.WindowsSamAccountName);
                        SPUser user = web.EnsureUser(userClaim.ToEncodedString());

                        string eventType = alertNode.Attributes["event"].Value;
                        SPEventType spEventType = (SPEventType)Enum.Parse(typeof(SPEventType), eventType);

                        string eventFrequency = alertNode.Attributes["frequency"].Value;
                        SPAlertFrequency spAlertFrequency = (SPAlertFrequency)Enum.Parse(typeof(SPAlertFrequency), eventFrequency);

                        string type = alertNode.Attributes["type"].Value;
                        SPAlertType spAlertType = (SPAlertType)Enum.Parse(typeof(SPAlertType), type);

                        SPListItem item = null;
                        if(list.BaseType == SPBaseType.DocumentLibrary)
                        {
                            SPFile file = web.GetFile(itemUrl);
                            item = file.Item;
                        }
                        else
                        {
                            item = list.GetItemById(itemId);
                        }

                        SPAlert newAlert = user.Alerts.Add();

                        newAlert.Title = alertTitle;
                        newAlert.AlertType = spAlertType;
                        newAlert.Item = item;
                        newAlert.DeliveryChannels = SPAlertDeliveryChannels.Email;
                        newAlert.EventType = spEventType;
                        newAlert.AlertFrequency = spAlertFrequency;
                        newAlert.Status = SPAlertStatus.On;
                        newAlert.Update(false);
                        output.Append(string.Format("Complete" + Environment.NewLine));
                    }
                    catch (Exception ex) { output.Append(string.Format("error: {0,20}" + Environment.NewLine, ex.Message)); }
                }
            }
        }
    }
}
