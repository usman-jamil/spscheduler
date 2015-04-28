using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Xml.Serialization;
using Microsoft.SharePoint;
using Emirates.Scheduler.SP2007.Tools.Structure;

namespace Emirates.Scheduler.SP2007.Tools
{
    public class SiteStructure : iTool
    {
        class subsite
        {
            public string source;
            public string target;

            public subsite()
            {
                source = string.Empty;
                target = string.Empty;
            }
        }

        class config
        {
            public string errorFile;
            public bool allWebs;

            public List<subsite> subSites;

            public config()
            {
                errorFile = string.Empty;
                allWebs = false;
                subSites = new List<subsite>();
            }

            public void ReadConfig(string inputXml)
            {
                XmlDocument xmDoc = new XmlDocument();
                xmDoc.Load(inputXml);

                XmlNode rootNode = xmDoc.SelectSingleNode("Sites");
                errorFile = rootNode.Attributes["ErrorFile"].Value;
                allWebs = Convert.ToBoolean(rootNode.Attributes["AllWebs"].Value);

                XmlNodeList sites = xmDoc.SelectNodes("//Site");
                foreach (XmlNode node in sites)
                {
                    string source = node.Attributes["Source"].Value;
                    string target = node.Attributes["Target"].Value;

                    subsite site = new subsite();
                    site.source = source;
                    site.target = target;

                    subSites.Add(site);
                }
            }
        }

        private sitestructure structure;

        public SiteStructure()
        {
            structure = new sitestructure();
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);

            config siteConfig = new config();
            siteConfig.ReadConfig(job.DownloadAttachment());

            List<site> sites = new List<site>();
            foreach (subsite subSite in siteConfig.subSites)
            {
                Console.WriteLine(subSite.source);
                using (SPWeb web = new SPSite(subSite.source).OpenWeb())
                {
                    site site = new site(subSite.source, subSite.target);
                    structure.sites.Add(site);
                    SPWebCollection subsites = siteConfig.allWebs ? web.Site.AllWebs : web.Webs;

                    foreach (SPWeb spWeb in subsites)
                    {
                        string oldUrl = spWeb.Url;
                        string siteRelativeUrl = oldUrl.Replace(subSite.source, string.Empty);

                        siteRelativeUrl = siteRelativeUrl.StartsWith("/") ? siteRelativeUrl : (string.IsNullOrEmpty(siteRelativeUrl) ? siteRelativeUrl : siteRelativeUrl.Substring(1));
                        string newUrl = subSite.target + siteRelativeUrl;

                        site newSite = new site(oldUrl, newUrl);
                        structure.sites.Add(newSite);
                        RecursiveWebCheck(spWeb, subSite.source, subSite.target);

                        spWeb.Dispose();
                    }
                }
            }

            XmlSerializer serializer = new XmlSerializer(typeof(sitestructure));
            string tmpFile = Scheduler.Instance.CreateTmpFile();
            using (TextWriter stream = new StreamWriter(tmpFile))
            {
                using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = true }))
                {
                    writer.WriteStartDocument();
                    serializer.Serialize(writer, structure);
                    writer.WriteEndDocument();
                    writer.Flush();
                }
            }

            result.AddFile(tmpFile);

            return result;
        }

        private void RecursiveWebCheck(SPWeb oSPWeb, string source, string target)
        {
            foreach (SPWeb web in oSPWeb.Webs)
            {
                string oldUrl = web.Url;
                string siteRelativeUrl = oldUrl.Replace(source, string.Empty);

                siteRelativeUrl = siteRelativeUrl.StartsWith("/") ? siteRelativeUrl : (string.IsNullOrEmpty(siteRelativeUrl) ? siteRelativeUrl : siteRelativeUrl.Substring(1));
                string newUrl = target + siteRelativeUrl;

                site newSite = new site(oldUrl, newUrl);
                structure.sites.Add(newSite);

                if (web.Webs.Count > 0)
                {
                    RecursiveWebCheck(web, source, target);
                }

                web.Dispose();
            }

        }
    }
}
