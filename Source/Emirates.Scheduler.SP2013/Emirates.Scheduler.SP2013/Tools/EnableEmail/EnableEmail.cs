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
    public class EnableEmail : iTool
    {
        StringBuilder output = null;

        public EnableEmail()
        {
            output = new StringBuilder();
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);
            string inputXml = job.DownloadAttachment();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputXml);

            XmlNodeList siteNodes = xmlDoc.SelectNodes("listreport/site");
            foreach (XmlNode siteNode in siteNodes)
            {
                string sourceUrl = siteNode.Attributes["source"].Value;
                string url = siteNode.Attributes["target"].Value;

                output.Append(string.Format("updating web: {0}" + Environment.NewLine, url));
                CheckLists(siteNode, url);
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

        private void CheckLists(XmlNode siteNode, string url)
        {
            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                XmlNodeList listNodes = siteNode.SelectNodes("list[@template='107'] | list[@template='1100']");
                foreach (XmlNode listNode in listNodes)
                {
                    string listTitle = listNode.Attributes["list"].Value;
                    output.Append(listTitle + Environment.NewLine);
                    try
                    {
                        SPList list = web.Lists[listTitle];
                        output.Append(string.Format("enabling list: {0}" + Environment.NewLine, listTitle));
                        list.EnableAssignToEmail = true;
                        list.Update();
                    }
                    catch { output.Append(string.Format("error: {0,20}" + Environment.NewLine, listTitle)); }
                }
            }
        }
    }
}
