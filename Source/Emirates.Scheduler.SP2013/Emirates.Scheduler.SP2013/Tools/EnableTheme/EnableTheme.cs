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
    public class EnableTheme : iTool
    {
        StringBuilder output = null;

        public EnableTheme()
        {
            output = new StringBuilder();
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);
            string inputXml = job.DownloadAttachment();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputXml);

            XmlNodeList siteNodes = xmlDoc.SelectNodes("listreport/site[@theme='GroupWorld']");
            foreach (XmlNode siteNode in siteNodes)
            {
                string sourceUrl = siteNode.Attributes["source"].Value;
                string url = siteNode.Attributes["target"].Value;

                output.Append(string.Format("updating web: {0}" + Environment.NewLine, url));
                CheckWeb(url);
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

        private void CheckWeb(string url)
        {
            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                try
                {
                    web.AllowUnsafeUpdates = true;
                    //web.ApplyTheme("Orange");
      //              string fontSchemeUrl = web.ServerRelativeUrl + "/_catalogs/theme/15/fontscheme003.spfont"
      //$themeurl = $SPWeb.ServerRelativeUrl + "/_catalogs/theme/15/palette005.spcolor"
      //$imageUrl = $SPWeb.ServerRelativeUrl + "/_layouts/15/images/image_bg005.jpg"
                    web.Update();
                    web.AllowUnsafeUpdates = false;
                }
                catch { output.Append(string.Format("error: {0,20}" + Environment.NewLine, url)); }
            }
        }
    }
}
