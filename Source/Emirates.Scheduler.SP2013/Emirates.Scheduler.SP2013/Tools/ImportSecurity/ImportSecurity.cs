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
    public class ImportSecurity : iTool
    {
        StringBuilder output = null;

        Result iTool.Execute(Job job)
        {
            output = new StringBuilder();

            Result result = new Result(job.Id);
            string inputXml = job.DownloadAttachment();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputXml);

            XmlNodeList siteNodes = xmlDoc.SelectNodes("security/site");
            foreach (XmlNode siteNode in siteNodes)
            {
                string sourceUrl = siteNode.Attributes["source"].Value;
                string url = siteNode.Attributes["target"].Value;

                output.Append(string.Format("updating web: {0}", url));
                CheckLists(siteNode, url);
                CheckFolders(siteNode, url);
                CheckFiles(siteNode, url);
            }

            string tmpFile = Scheduler.Instance.CreateTmpFile();

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
                XmlNodeList listNodes = siteNode.SelectNodes("folder[@list='true']");
                foreach (XmlNode listNode in listNodes)
                {
                    string listTitle = listNode.Attributes["folder"].Value;
                    Console.WriteLine(listTitle);
                    try
                    {
                        SPList list = web.Lists[listTitle];
                        output.Append(string.Format("checking list: {0}", listTitle));
                        SPRoleAssignmentCollection roleAssignments = list.RoleAssignments;
                        if (list.HasUniqueRoleAssignments)
                        {
                            XmlNodeList principalGroupNodes = listNode.SelectNodes("principal[@Group='true']");
                            CheckGroups(web, list, principalGroupNodes);

                            XmlNodeList principalUserNodes = listNode.SelectNodes("principal[@Group='false']");
                            CheckUsers(web, list, principalUserNodes);
                        }
                        else
                        {
                            output.Append(string.Format("target list: {0,20}, is inheriting permissions", listTitle));
                        }
                    }
                    catch { output.Append(string.Format("list missing: {0,20}", listTitle)); }
                }
            }
        }

        private void CheckFolders(XmlNode siteNode, string url)
        {
            bool containsIgnoreInheritance = ContainsAttribute("ignoreinheritance", siteNode);
            bool containsImportOptions = ContainsAttribute("import", siteNode);

            bool ignoreInheritance = containsIgnoreInheritance ?
                Boolean.Parse(siteNode.Attributes["ignoreinheritance"].Value) :
                true;

            ImportOptions importOptions = containsImportOptions ?
                (ImportOptions)Enum.Parse(typeof(ImportOptions), siteNode.Attributes["import"].Value) :
                ImportOptions.All;

            if (importOptions == ImportOptions.Files)
                return;

            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                XmlNodeList folderNodes = siteNode.SelectNodes("folder[@list='false']");
                foreach (XmlNode folderNode in folderNodes)
                {
                    string folderTitle = folderNode.Attributes["folder"].Value;
                    string folderUrl = folderNode.Attributes["url"].Value;
                    output.Append(string.Format("checking folder: {0}", folderTitle));
                    try
                    {
                        SPFolder folder = web.GetFolder(folderUrl);
                        SPListItem item = folder.Item;

                        bool breakInheritance = !item.HasUniqueRoleAssignments && !ignoreInheritance;
                        bool applyPermissions = item.HasUniqueRoleAssignments || breakInheritance;

                        if (applyPermissions)
                        {
                            if (breakInheritance)
                            {
                                output.Append(string.Format("Breaking Inheritance!"));
                                item.BreakRoleInheritance(false, false);
                            }

                            XmlNodeList principalGroupNodes = folderNode.SelectNodes("principal[@Group='true']");
                            CheckGroups(web, item, principalGroupNodes);

                            XmlNodeList principalUserNodes = folderNode.SelectNodes("principal[@Group='false']");
                            CheckUsers(web, item, principalUserNodes);
                        }
                        else
                        {
                            output.Append(string.Format("target folder: {0,20}, is inheriting permissions", folderTitle));
                        }
                    }
                    catch { output.Append(string.Format("folder missing: {0,20}", folderTitle)); }
                }
            }
        }

        private void CheckFiles(XmlNode siteNode, string url)
        {
            bool containsIgnoreInheritance = ContainsAttribute("ignoreinheritance", siteNode);
            bool containsImportOptions = ContainsAttribute("import", siteNode);

            bool ignoreInheritance = containsIgnoreInheritance ?
                Boolean.Parse(siteNode.Attributes["ignoreinheritance"].Value) :
                true;

            ImportOptions importOptions = containsImportOptions ?
                (ImportOptions)Enum.Parse(typeof(ImportOptions), siteNode.Attributes["import"].Value) :
                ImportOptions.All;

            if (importOptions == ImportOptions.Folders)
                return;

            using (SPWeb web = new SPSite(url).OpenWeb())
            {
                XmlNodeList fileNodes = siteNode.SelectNodes("file");

                foreach (XmlNode fileNode in fileNodes)
                {
                    string fileTitle = fileNode.Attributes["file"].Value;
                    string fileUrl = fileNode.Attributes["url"].Value;
                    Console.WriteLine("checking file: {0}", fileTitle);
                    try
                    {
                        SPFile file = web.GetFile(fileUrl);
                        SPListItem item = file.Item;

                        bool breakInheritance = !item.HasUniqueRoleAssignments && !ignoreInheritance;
                        bool applyPermissions = item.HasUniqueRoleAssignments || breakInheritance;

                        if (applyPermissions)
                        {
                            if (breakInheritance)
                            {
                                output.Append(string.Format("Breaking Inheritance!"));
                                item.BreakRoleInheritance(false, false);
                            }

                            XmlNodeList principalGroupNodes = fileNode.SelectNodes("principal[@Group='true']");
                            CheckGroups(web, item, principalGroupNodes);

                            XmlNodeList principalUserNodes = fileNode.SelectNodes("principal[@Group='false']");
                            CheckUsers(web, item, principalUserNodes);
                        }
                        else
                        {
                            output.Append(string.Format("target file: {0,20}, is inheriting permissions", fileTitle));
                        }
                    }
                    catch { output.Append(string.Format("file missing: {0,20}", fileTitle)); }
                }
            }
        }

        private void CheckGroups(SPWeb web, SPList list, XmlNodeList principalGroupNodes)
        {
            SPRoleAssignmentCollection roleAssignments = list.RoleAssignments;
            bool updated = UpdateGroupPermissions(web, roleAssignments, principalGroupNodes);

            if (updated)
                list.Update();
        }

        private void CheckGroups(SPWeb web, SPListItem item, XmlNodeList principalGroupNodes)
        {
            SPRoleAssignmentCollection roleAssignments = item.RoleAssignments;
            bool updated = UpdateGroupPermissions(web, roleAssignments, principalGroupNodes);

            if (updated)
                item.SystemUpdate();
        }

        private bool UpdateGroupPermissions(SPWeb web, SPRoleAssignmentCollection roleAssignments, XmlNodeList principalGroupNodes)
        {
            bool dirty = false;
            foreach (XmlNode principalGroupNode in principalGroupNodes)
            {
                string groupName = principalGroupNode.Attributes["name"].Value;
                try
                {
                    SPGroup group = web.SiteGroups.GetByName(groupName);
                    if (group == null)
                        throw new Exception();

                    SPPrincipal groupPrincipal = (SPPrincipal)group;

                    try
                    {
                        SPRoleAssignment roleAssignment = roleAssignments.GetAssignmentByPrincipal(groupPrincipal);
                        if (roleAssignment == null)
                            throw new Exception();

                        UpdatePrincipal(web, principalGroupNode, groupName, roleAssignment);
                    }
                    catch
                    {
                        Console.Write("permissins missing for: {0,20}, adding new...", groupName);
                        SPRoleAssignment roleAssignmentNew = new SPRoleAssignment(groupPrincipal);
                        XmlNodeList roleNodes = principalGroupNode.SelectNodes("role");
                        foreach (XmlNode roleNode in roleNodes)
                        {
                            string roleName = roleNode.Attributes["name"].Value;
                            if (roleName.ToLower().Equals("limited access"))
                                roleName = "Limited User";

                            SPRoleDefinition role = web.RoleDefinitions[roleName];
                            roleAssignmentNew.RoleDefinitionBindings.Add(role);
                        }
                        roleAssignments.Add(roleAssignmentNew);
                        Console.WriteLine("completed");
                        dirty = true;
                    }
                }
                catch { Console.WriteLine("group not found: {0,20}", groupName); }
            }

            return dirty;
        }

        private void CheckUsers(SPWeb web, SPList list, XmlNodeList principalGroupNodes)
        {
            SPRoleAssignmentCollection roleAssignments = list.RoleAssignments;
            bool updated = UpdateUserPermissions(web, roleAssignments, principalGroupNodes);

            if (updated)
                list.Update();
        }

        private void CheckUsers(SPWeb web, SPListItem item, XmlNodeList principalGroupNodes)
        {
            SPRoleAssignmentCollection roleAssignments = item.RoleAssignments;
            bool updated = UpdateUserPermissions(web, roleAssignments, principalGroupNodes);

            if (updated)
                item.SystemUpdate();
        }

        private bool UpdateUserPermissions(SPWeb web, SPRoleAssignmentCollection roleAssignments, XmlNodeList principalUserNodes)
        {
            bool dirty = false;
            foreach (XmlNode principalUserNode in principalUserNodes)
            {
                string loginName = principalUserNode.Attributes["login"].Value;
                string userName = principalUserNode.Attributes["name"].Value;
                try
                {
                    SPClaimProviderManager cpm = SPClaimProviderManager.Local;
                    SPClaim userClaim = cpm.ConvertIdentifierToClaim(loginName, SPIdentifierTypes.WindowsSamAccountName);

                    SPUser user = web.EnsureUser(userClaim.ToEncodedString());
                    if (user == null)
                        throw new Exception();

                    SPPrincipal userPrincipal = (SPPrincipal)user;

                    try
                    {
                        SPRoleAssignment roleAssignment = roleAssignments.GetAssignmentByPrincipal(userPrincipal);
                        if (roleAssignment == null)
                            throw new Exception();

                        UpdatePrincipal(web, principalUserNode,
                            string.Format("{0,20}, '{1,15}'", userName, loginName),
                            roleAssignment);
                    }
                    catch
                    {
                        Console.Write("permissins missing for user: {0,20} with login: {1,15}, adding new...", userName, loginName);
                        SPRoleAssignment roleAssignmentNew = new SPRoleAssignment(userPrincipal);
                        XmlNodeList roleNodes = principalUserNode.SelectNodes("role");
                        foreach (XmlNode roleNode in roleNodes)
                        {
                            string roleName = roleNode.Attributes["name"].Value;
                            if (roleName.ToLower().Equals("limited access"))
                                roleName = "Limited User";

                            SPRoleDefinition role = web.RoleDefinitions[roleName];
                            roleAssignmentNew.RoleDefinitionBindings.Add(role);
                        }
                        roleAssignments.Add(roleAssignmentNew);
                        Console.WriteLine("completed");
                        dirty = true;
                    }
                }
                catch { Console.WriteLine("user not found: {0,20} with login: {1,15}", userName, loginName); }
            }

            return dirty;
        }

        private void UpdatePrincipal(SPWeb web, XmlNode principalGroupNode, string principalName, SPRoleAssignment roleAssignment)
        {
            XmlNodeList roleNodes = principalGroupNode.SelectNodes("role");
            foreach (XmlNode roleNode in roleNodes)
            {
                string roleName = roleNode.Attributes["name"].Value;
                if (roleName.ToLower().Equals("limited access"))
                    roleName = "Limited User";

                bool found = false;
                foreach (SPRoleDefinition roleDefinition in roleAssignment.RoleDefinitionBindings)
                {
                    if (roleDefinition.Name.ToLower().Equals(roleName.ToLower()))
                    {
                        found = true;
                    }
                }

                if (!found)
                {
                    Console.Write("role: {0,15} missing for principal: {1,20}, adding new...",
                        roleName,
                        principalName);

                    SPRoleDefinition role = web.RoleDefinitions[roleName];
                    roleAssignment.RoleDefinitionBindings.Add(role);
                    roleAssignment.Update();
                    Console.WriteLine("completed");
                }
            }
        }
    }
}
