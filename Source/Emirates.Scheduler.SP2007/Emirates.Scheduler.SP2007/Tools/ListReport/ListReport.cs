using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Xml.Serialization;
using Microsoft.SharePoint;
using Emirates.Scheduler.SP2007.Tools.Lists;

namespace Emirates.Scheduler.SP2007.Tools
{
    public class ListReport : iTool
    {
        class comparisonsite
        {
            public string source;
            public string target;

            public comparisonsite()
            {
                source = string.Empty;
                target = string.Empty;
            }
        }

        class config
        {
            public List<string> ignoreList;

            public List<comparisonsite> comparisonSites;

            public config()
            {
                comparisonSites = new List<comparisonsite>();
            }

            public void ReadConfig(string inputXml)
            {
                XmlDocument xmDoc = new XmlDocument();
                xmDoc.Load(inputXml);

                XmlNode rootNode = xmDoc.SelectSingleNode("Sites");
                string[] ignoreString = rootNode.Attributes["Ignore-List"].Value.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                ignoreList = new List<string>(ignoreString);

                XmlNodeList sites = xmDoc.SelectNodes("//Site");
                foreach (XmlNode node in sites)
                {
                    string source = node.Attributes["Source"].Value;
                    string target = node.Attributes["Target"].Value;

                    comparisonsite site = new comparisonsite();
                    site.source = source;
                    site.target = target;

                    comparisonSites.Add(site);
                }
            }
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);

            config permConfig = new config();
            listreport security = new listreport();
            permConfig.ReadConfig(job.DownloadAttachment());

            List<site> sites = new List<site>();
            foreach (comparisonsite compSite in permConfig.comparisonSites)
            {
                Console.WriteLine(compSite.source);
                using (SPWeb web = new SPSite(compSite.source).OpenWeb())
                {
                    int listCount = 0;
                    site site = new site(compSite.source, compSite.target);
                    List<list> folders = new List<list>();
                    SPListCollection siteLists = web.Lists;

                    foreach (SPList list in siteLists)
                    {
                        if (!permConfig.ignoreList.Contains(list.RootFolder.Name.ToLower()))
                        {
                            listCount++;
                            list listFolder = AddList(list);

                            Helper helper = Helper.Instance;

                            string updatedUrl = helper.MapServerRelativeUrl(listFolder.serverRelativeUrl,
                                compSite.source,
                                compSite.target);
                            listFolder.serverRelativeUrl = updatedUrl;

                            site.lists.Add(listFolder);
                        }
                    }

                    site.ListCount = listCount;
                    site.theme = web.Theme;
                    security.sites.Add(site);
                }
            }

            XmlSerializer serializer = new XmlSerializer(typeof(listreport));
            string tmpFile = Scheduler.Instance.CreateTmpFile();
            using (TextWriter stream = new StreamWriter(tmpFile))
            {
                using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = true }))
                {
                    writer.WriteStartDocument();
                    serializer.Serialize(writer, security);
                    writer.WriteEndDocument();
                    writer.Flush();
                }
            }

            result.AddFile(tmpFile);

            return result;
        }

        private list AddList(SPList list)
        {
            list newList = new list();
            newList.folderName = list.Title;
            newList.folderName = newList.folderName.Replace("\v", " ");
            newList.serverRelativeUrl = list.RootFolder.ServerRelativeUrl;
            newList.template = (int)list.BaseTemplate;
            newList.templateType = list.BaseTemplate.ToString();
            newList.baseType = list.BaseType.ToString();
            newList.emailAlias = list.EmailAlias;
            newList.enableAssignToEmail = list.ContentTypesEnabled;
            newList.workflows = list.WorkflowAssociations.Count;

            return newList;
        }
    }
}
