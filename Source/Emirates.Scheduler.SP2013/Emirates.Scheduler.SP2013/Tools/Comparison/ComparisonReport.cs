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
    public class ComparisonReport : iTool
    {
        StringBuilder output = null;

        public ComparisonReport()
        {
            output = new StringBuilder();
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);
            string inputXml = job.DownloadAttachment();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputXml);

            XmlNodeList siteNodes = xmlDoc.SelectNodes("comparison/site");
            foreach (XmlNode siteNode in siteNodes)
            {
                string sourceUrl = siteNode.Attributes["source"].Value;
                string url = siteNode.Attributes["target"].Value;

                output.Append(string.Format("comparing web: {0}" + Environment.NewLine, url));
                try
                {
                    if (CheckWeb(siteNode, url))
                    {
                        CheckLists(siteNode, url);
                        CheckFolders(siteNode, url);
                    }
                }
                catch { output.Append(string.Format("web missing: {0,20}" + Environment.NewLine, url)); }
            }

            string tmpFile = Scheduler.Instance.CreateTmpFile();

            System.IO.File.WriteAllText(tmpFile, output.ToString());

            result.AddFile(tmpFile);
            return result;
        }

        private bool ContainsAttribute(string attr, XmlNode node)
        {
            bool found = false;

            try
            {
                XmlAttribute attribute = node.Attributes[attr];
                found = attribute != null;
            }
            catch { }

            return found;
        }

        private bool CheckWeb(XmlNode siteNode, string url)
        {
            bool valid = false;

            try
            {
                using (SPWeb web = new SPSite(url).OpenWeb())
                {
                    valid = web.Exists;
                }
            }
            catch { }

            return valid;
        }

        private void CheckLists(XmlNode siteNode, string url)
        {
            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                XmlNodeList listNodes = siteNode.SelectNodes("folder[@list='true']");
                foreach (XmlNode listNode in listNodes)
                {
                    string listTitle = listNode.Attributes["folder"].Value;
                    int itemCount = Convert.ToInt32(listNode.Attributes["count"].Value);

                    try
                    {
                        SPList list = web.Lists[listTitle];
                        int listItemCount = list.ItemCount;
                        if (listItemCount != itemCount)
                        {
                            output.Append(string.Format("list: {0}, count differs. Source has {1} items and target has {2}: " + Environment.NewLine, 
                                listTitle,
                                itemCount,
                                listItemCount));
                        }
                    }
                    catch { output.Append(string.Format("list missing: {0,20}" + Environment.NewLine, listTitle)); }
                }
            }
        }

        private void CheckFolders(XmlNode siteNode, string url)
        {
            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                XmlNodeList folderNodes = siteNode.SelectNodes("folder[@list='false']");
                foreach (XmlNode folderNode in folderNodes)
                {
                    string folderTitle = folderNode.Attributes["folder"].Value;
                    string folderUrl = folderNode.Attributes["url"].Value;
                    int itemCount = Convert.ToInt32(folderNode.Attributes["count"].Value);

                    try
                    {
                        SPFolder folder = web.GetFolder(folderUrl);
                        SPListItem item = folder.Item;
                        int fileCount = folder.Files.Count;

                        if (itemCount != fileCount)
                        {
                            output.Append(string.Format("folder: {0}, count differs. Source has {1} files and target has {2}: " + Environment.NewLine,
                                folderUrl,
                                itemCount,
                                fileCount));
                        }
                    }
                    catch { output.Append(string.Format("folder missing: {0,20}" + Environment.NewLine, folderUrl)); }
                }
            }
        }
    }
}
