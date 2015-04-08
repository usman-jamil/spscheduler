using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Xml.Serialization;

namespace Emirates.Scheduler.SP2007.Tools.Security
{
    [Serializable]
    public class security
    {
        [XmlElement("site")]
        public List<site> sites;

        public security()
        {
            sites = new List<site>();
        }

        public void AddSite(site newSite)
        {
            sites.Add(newSite);
        }
    }

    [Serializable]
    public class site
    {
        [XmlAttribute("source")]
        public string source;
        [XmlAttribute("target")]
        public string target;
        [XmlAttribute("import")]
        public string import;
        [XmlAttribute("ignoreinheritance")]
        public bool ignoreInheritance;

        [XmlElement("folder")]
        public List<folder> folders;

        [XmlElement("file")]
        public List<file> files;

        public site()
        {
        }

        public site(string source, string target)
        {
            this.source = source;
            this.target = target;
            import = "All";
            ignoreInheritance = true;
            folders = new List<folder>();
            files = new List<file>();
        }

        public void AddFolder(string folderName, string serverRelativeUrl, bool isSharePointList)
        {
            folders.Add(new folder(folderName, serverRelativeUrl, isSharePointList));
        }

        public void AddFolder(string folderName, string serverRelativeUrl)
        {
            folders.Add(new folder(folderName, serverRelativeUrl, false));
        }

        public void AddFile(string fileName, string serverRelativeUrl)
        {
            files.Add(new file(fileName, serverRelativeUrl));
        }
    }

    [Serializable]
    public class folder
    {
        [XmlAttribute("folder")]
        public string folderName;
        [XmlAttribute("url")]
        public string serverRelativeUrl;
        [XmlAttribute("list")]
        public bool isSharePointList;

        [XmlElement("principal")]
        public List<principal> principals;

        public folder()
        {
            folderName = string.Empty;
            serverRelativeUrl = string.Empty;
            isSharePointList = false;
            principals = new List<principal>();
        }

        public folder(string folderName, string serverRelativeUrl, bool isSharePointList)
        {
            this.folderName = folderName;
            this.serverRelativeUrl = serverRelativeUrl;
            this.isSharePointList = isSharePointList;
        }

        public void AddPrincipal(string login, string name, bool isGroup, SPRoleDefinitionBindingCollection bindings)
        {
            principals.Add(new principal(login, name, isGroup, bindings));
        }
    }

    [Serializable]
    public class file
    {
        [XmlAttribute("file")]
        public string fileName;
        [XmlAttribute("url")]
        public string serverRelativeUrl;

        [XmlElement("principal")]
        public List<principal> principals;

        public file()
        {
            fileName = string.Empty;
            serverRelativeUrl = string.Empty;
            principals = new List<principal>();
        }

        public file(string fileName, string serverRelativeUrl)
        {
            this.fileName = fileName;
            this.serverRelativeUrl = serverRelativeUrl;
        }

        public void AddPrincipal(string login, string name, bool isGroup, SPRoleDefinitionBindingCollection bindings)
        {
            principals.Add(new principal(login, name, isGroup, bindings));
        }
    }

    [Serializable]
    public class principal
    {
        [XmlAttribute("login")]
        public string login;
        [XmlAttribute("name")]
        public string name;
        [XmlAttribute("Group")]
        public bool isGroup;

        [XmlElement("role")]
        public List<role> roles;

        public principal()
        {
            roles = new List<role>();
        }

        public principal(string login, string name, bool isGroup, SPRoleDefinitionBindingCollection bindings)
        {
            this.login = login;
            this.name = name;
            this.isGroup = isGroup;
            if (roles == null)
            {
                roles = new List<role>();
            }

            foreach (SPRoleDefinition binding in bindings)
            {
                roles.Add(new role(binding.Name));
            }
        }
    }

    [Serializable]
    public class role
    {
        [XmlAttribute("name")]
        public string name;

        public role()
        {
        }

        public role(string name)
        {
            this.name = name;
        }
    }
}