using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Xml.Serialization;

namespace Emirates.Scheduler.SP2007.Tools.Lists
{
    [Serializable]
    public class listreport
    {
        [XmlElement("site")]
        public List<site> sites;

        public listreport()
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
        [XmlAttribute("listcount")]
        public int listCount;
        [XmlAttribute("theme")]
        public string theme;

        [XmlElement("list")]
        public List<list> lists;

        public site()
        {
        }

        public site(string source, string target)
            : this(source, target, 0, string.Empty)
        { }

        public site(string source, string target, int listCount, string theme)
        {
            this.source = source;
            this.target = target;
            this.listCount = listCount;
            this.theme = theme;
            lists = new List<list>();
        }

        public void AddFolder(string folderName, 
            string serverRelativeUrl, 
            int template, 
            string templateType, 
            string baseType, 
            string email,
            bool enableAssignToEmail,
            int workflows)
        {
            lists.Add(new list(folderName,
                        serverRelativeUrl, 
                        template, 
                        templateType, 
                        baseType, 
                        email,
                        enableAssignToEmail,
                        workflows));
        }

        public int ListCount
        {
            set
            {
                this.listCount = value;
            }
        }
    }

    [Serializable]
    public class list
    {
        [XmlAttribute("list")]
        public string folderName;
        [XmlAttribute("url")]
        public string serverRelativeUrl;
        [XmlAttribute("template")]
        public int template;
        [XmlAttribute("templatetype")]
        public string templateType;
        [XmlAttribute("basetype")]
        public string baseType;
        [XmlAttribute("emailalias")]
        public string emailAlias;
        [XmlAttribute("enableassigntoemail")]
        public bool enableAssignToEmail;
        [XmlAttribute("workflows")]
        public int workflows;

        public list()
        {
            folderName = string.Empty;
            serverRelativeUrl = string.Empty;
            template = 0;
            templateType = string.Empty;
            baseType = string.Empty;
            emailAlias = string.Empty;
        }

        public list(string folderName, 
            string serverRelativeUrl, 
            int template, 
            string templateType, 
            string baseType,
            string emailAlias,
            bool enableAssignToEmail,
            int workflows)
        {
            this.folderName = folderName;
            this.serverRelativeUrl = serverRelativeUrl;
            this.template = template;
            this.templateType = templateType;
            this.baseType = baseType;
            this.emailAlias = emailAlias;
            this.enableAssignToEmail = enableAssignToEmail;
            this.workflows = workflows;
        }
    }
}