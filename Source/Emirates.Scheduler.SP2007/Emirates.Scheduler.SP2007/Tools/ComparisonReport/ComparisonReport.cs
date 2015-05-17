using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Xml.Serialization;
using Microsoft.SharePoint;
using Emirates.Scheduler.SP2007.Tools.Report;

namespace Emirates.Scheduler.SP2007.Tools
{
    public class ComparisonReport : iTool
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
            public string errorFile;
            public List<string> ignoreList;

            public List<comparisonsite> comparisonSites;

            public config()
            {
                errorFile = string.Empty;
                comparisonSites = new List<comparisonsite>();
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
            comparison security = new comparison();
            Helper helper = Helper.Instance;
            permConfig.ReadConfig(job.DownloadAttachment());

            List<site> sites = new List<site>();
            foreach (comparisonsite compSite in permConfig.comparisonSites)
            {
                Console.WriteLine(compSite.source);
                using (SPWeb web = new SPSite(compSite.source).OpenWeb())
                {
                    int listCount = 0;
                    site site = new site(compSite.source, compSite.target);
                    List<folder> folders = new List<folder>();
                    SPListCollection siteLists = web.Lists;

                    foreach (SPList list in siteLists)
                    {
                        if (!permConfig.ignoreList.Contains(list.RootFolder.Name.ToLower()))
                        {
                            listCount++;
                            folder listFolder = AddFolder(list);

                            listFolder.serverRelativeUrl = helper.MapServerRelativeUrl(listFolder.serverRelativeUrl,
                                compSite.source,
                                compSite.target);

                            site.folders.Add(listFolder);

                            SPQuery query = new SPQuery();
                            query.Query = @"
                                <Where>
                                    <BeginsWith>
                                        <FieldRef Name='ContentTypeId' />
                                        <Value Type='ContentTypeId'>0x0120</Value>
                                    </BeginsWith>
                                </Where>";
                            query.ViewAttributes = "Scope='RecursiveAll'";
                            SPListItemCollection items = list.GetItems(query);

                            foreach (SPListItem item in items)
                            {
                                folder folder = AddFolder(item.Folder, false, item.Folder.Files.Count);

                                string updatedUrl = helper.MapServerRelativeUrl(folder.serverRelativeUrl,
                                    compSite.source,
                                    compSite.target);
                                folder.serverRelativeUrl = updatedUrl;

                                site.folders.Add(folder);
                            }
                        }
                    }

                    site.ListCount = listCount;
                    security.sites.Add(site);
                }
            }

            XmlSerializer serializer = new XmlSerializer(typeof(comparison));
            string tmpFile = Scheduler.Instance.CreateTmpFile();
            using (TextWriter stream = new StreamWriter(tmpFile))
            {
                using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = true, CheckCharacters = true }))
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

        private folder AddFolder(SPList list)
        {
            return AddFolder(list.RootFolder, true, list.ItemCount);
        }

        private folder AddFolder(SPFolder spFolder, bool isList, int itemCount)
        {
            Helper helper = Helper.Instance;

            folder folder = new folder();
            folder.folderName = isList ? spFolder.ParentWeb.Lists[spFolder.ParentListId].Title : spFolder.Name;
            folder.folderName = folder.folderName.Replace("\v", " ");
            folder.serverRelativeUrl = spFolder.ServerRelativeUrl;
            folder.isSharePointList = isList;
            folder.itemCount = itemCount;

            return folder;
        }
    }
}
