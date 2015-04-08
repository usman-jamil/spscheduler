using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Xml.Serialization;
using Microsoft.SharePoint;
using Emirates.Scheduler.SP2007.Tools.Security;

namespace Emirates.Scheduler.SP2007.Tools
{
    public class ExportSecurity : iTool
    {
        class permissionsite
        {
            public string source;
            public string target;
            public bool ignoreFiles;

            public permissionsite()
            {
                source = string.Empty;
                target = string.Empty;
                ignoreFiles = false;
            }
        }

        class config
        {
            public string errorFile;
            public List<string> ignoreList;

            public List<permissionsite> permissionSites;

            public config()
            {
                errorFile = string.Empty;
                permissionSites = new List<permissionsite>();
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
                    string sIgnoreFiles = node.Attributes["IgnoreFiles"].Value;

                    permissionsite site = new permissionsite();
                    site.source = source;
                    site.target = target;
                    site.ignoreFiles = Boolean.Parse(sIgnoreFiles);

                    permissionSites.Add(site);
                }
            }
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);

            config permConfig = new config();
            security security = new security();
            permConfig.ReadConfig(job.DownloadAttachment());

            List<site> sites = new List<site>();
            foreach (permissionsite permSite in permConfig.permissionSites)
            {
                Console.WriteLine(permSite.source);
                using (SPWeb web = new SPSite(permSite.source).OpenWeb())
                {
                    site site = new site(permSite.source, permSite.target);
                    List<folder> folders = new List<folder>();
                    SPListCollection siteLists = web.Lists;

                    foreach (SPList list in siteLists)
                    {
                        if (!permConfig.ignoreList.Contains(list.RootFolder.Name.ToLower()))
                        {
                            if (list.HasUniqueRoleAssignments)
                            {
                                folder folder = AddFolder(list);

                                site.folders.Add(folder);
                            }

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
                                if (item.HasUniqueRoleAssignments)
                                {
                                    folder folder = AddFolder(item.Folder, item.RoleAssignments, false);

                                    site.folders.Add(folder);
                                }

                                if (!permSite.ignoreFiles)
                                {
                                    List<file> uniqueSPFiles = GetUniquePermissionFiles(list, item);

                                    foreach (file uniqueSPFile in uniqueSPFiles)
                                    {
                                        site.files.Add(uniqueSPFile);
                                    }
                                }
                            }
                        }
                    }

                    security.sites.Add(site);
                }
            }

            XmlSerializer serializer = new XmlSerializer(typeof(security));
            string tmpFile = Scheduler.Instance.CreateTmpFile();
            using (TextWriter stream = new StreamWriter(tmpFile))
            {
                using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = true }))
                {
                    writer.WriteStartDocument();
                    writer.WriteComment(@"You can control what to import on the target by setting the Import attribute on the <Sites> element e.g. Import=""All"" | Import=""Folders"" | Import=""Files""");
                    writer.WriteComment(@"You can disconnect permission inheritance (if wrongfully inheriting) by setting IgnoreInheritance=""false""");
                    serializer.Serialize(writer, security);
                    writer.WriteEndDocument();
                    writer.Flush();
                }
            }

            result.AddFile(tmpFile);

            return result;
        }

        private List<file> GetUniquePermissionFiles(SPList list, SPListItem item)
        {
            List<file> uniqueFiles = new List<file>();

            SPQuery filesQuery = new SPQuery();
            filesQuery.Query = @"<OrderBy><FieldRef Name='Created' Ascending='FALSE'/></OrderBy>";
            filesQuery.ViewAttributes = "Scope='FilesOnly'";
            filesQuery.RowLimit = 99999;
            filesQuery.Folder = item.Folder;
            SPListItemCollection spListItems = list.GetItems(filesQuery);

            foreach (SPListItem spListItem in spListItems)
            {
                if (spListItem.HasUniqueRoleAssignments)
                {
                    uniqueFiles.Add(AddFile(spListItem.File, spListItem.RoleAssignments));
                }
            }

            return uniqueFiles;
        }

        private folder AddFolder(SPList list)
        {
            return AddFolder(list.RootFolder, list.RoleAssignments, true);
        }

        private folder AddFolder(SPFolder spFolder, SPRoleAssignmentCollection roleAssignments, bool isList)
        {
            folder folder = new folder();
            folder.folderName = isList ? spFolder.ParentWeb.Lists[spFolder.ParentListId].Title : spFolder.Name;
            folder.serverRelativeUrl = spFolder.ServerRelativeUrl;
            folder.isSharePointList = isList;

            foreach (SPRoleAssignment roleAssignment in roleAssignments)
            {
                SPPrincipal principal = roleAssignment.Member;
                string principalLogin = (principal is SPUser) ? principal.ParentWeb.AllUsers.GetByID(principal.ID).LoginName : principal.ID.ToString();
                bool isGroup = !(principal is SPUser);

                folder.AddPrincipal(principalLogin,
                    principal.Name,
                    isGroup,
                    roleAssignment.RoleDefinitionBindings);
            }

            return folder;
        }

        private file AddFile(SPFile spFile, SPRoleAssignmentCollection roleAssignments)
        {
            file file = new file();
            file.fileName = spFile.Name;
            file.serverRelativeUrl = spFile.ServerRelativeUrl;

            foreach (SPRoleAssignment roleAssignment in roleAssignments)
            {
                SPPrincipal principal = roleAssignment.Member;
                string principalLogin = (principal is SPUser) ? principal.ParentWeb.AllUsers.GetByID(principal.ID).LoginName : principal.ID.ToString();
                bool isGroup = !(principal is SPUser);

                file.AddPrincipal(principalLogin,
                    principal.Name,
                    isGroup,
                    roleAssignment.RoleDefinitionBindings);
            }

            return file;
        }
    }
}
