using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Xml.Serialization;
using Emirates.Scheduler.SP2007.Tools.Alerts;
using Microsoft.SharePoint;

namespace Emirates.Scheduler.SP2007.Tools
{
    public class ExportAlerts : iTool
    {
        class alertsite
        {
            public string source;
            public string target;

            public alertsite()
            {
                source = string.Empty;
                target = string.Empty;
            }
        }

        class config
        {
            public string errorFile;
            public List<string> ignoreList;

            public List<alertsite> alertSites;

            public config()
            {
                errorFile = string.Empty;
                alertSites = new List<alertsite>();
            }

            public void ReadConfig(string inputXml)
            {
                XmlDocument xmDoc = new XmlDocument();
                xmDoc.Load(inputXml);

                XmlNode rootNode = xmDoc.SelectSingleNode("Sites");
                errorFile = rootNode.Attributes["ErrorFile"].Value;
                string[] ignoreString = rootNode.Attributes["Ignore-List"].Value.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                ignoreList = new List<string>(ignoreString);

                XmlNodeList sites = xmDoc.SelectNodes("//Site");
                foreach (XmlNode node in sites)
                {
                    string source = node.Attributes["Source"].Value;
                    string target = node.Attributes["Target"].Value;

                    alertsite site = new alertsite();
                    site.source = source;
                    site.target = target;

                    alertSites.Add(site);
                }
            }
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);

            config permConfig = new config();
            notifications siteNotifications = new notifications();
            permConfig.ReadConfig(job.DownloadAttachment());

            List<site> sites = new List<site>();
            foreach (alertsite permSite in permConfig.alertSites)
            {
                using (SPWeb web = new SPSite(permSite.source).OpenWeb())
                {
                    try
                    {
                        Console.WriteLine(web.Url);
                        site site = new site(permSite.source, permSite.target);
                        SPListCollection siteLists = web.Lists;
                        SPAlertCollection allAlerts = web.Alerts;
                        Console.WriteLine(allAlerts.Count);

                        foreach (SPAlert spAlert in allAlerts)
                        {
                            try
                            {
                                if (spAlert.AlertType == SPAlertType.List)
                                {
                                    site.AddAlert(spAlert.User.LoginName,
                                        spAlert.List.Title,
                                        spAlert.EventType.ToString(),
                                        spAlert.AlertFrequency.ToString(),
                                        spAlert.AlertType.ToString(),
                                        string.Empty,
                                        false);
                                }
                                else
                                {
                                    string url = (spAlert.List.BaseType == SPBaseType.DocumentLibrary) ?
                                        spAlert.Item.File.ServerRelativeUrl :
                                        spAlert.Item.Url;
                                    bool isFile = (spAlert.List.BaseType == SPBaseType.DocumentLibrary) ?
                                        true :
                                        false;
                                    site.AddAlert(spAlert.User.LoginName,
                                        spAlert.List.Title,
                                        spAlert.EventType.ToString(),
                                        spAlert.AlertFrequency.ToString(),
                                        spAlert.AlertType.ToString(),
                                        spAlert.ItemID,
                                        url,
                                        isFile);
                                }
                            }
                            catch (Exception ex) { Console.WriteLine(ex.Message); }
                        }

                        siteNotifications.sites.Add(site);
                    }
                    catch (Exception ex) { Console.WriteLine(ex.Message); }
                }
            }

            XmlSerializer serializer = new XmlSerializer(typeof(notifications));
            string tmpFile = Scheduler.Instance.CreateTmpFile();
            using (TextWriter writer = new StreamWriter(tmpFile))
            {
                serializer.Serialize(writer, siteNotifications);
            }

            result.AddFile(tmpFile);

            return result;
        }
    }
}
