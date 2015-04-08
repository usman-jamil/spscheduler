using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Xml.Serialization;

namespace Emirates.Scheduler.SP2007.Tools.Report
{
    [Serializable]
    public class comparison
    {
        [XmlElement("site")]
        public List<site> sites;

        public comparison()
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

        [XmlElement("folder")]
        public List<folder> folders;

        public site()
        {
        }

        public site(string source, string target)
            : this(source, target, 0)
        { }

        public site(string source, string target, int listCount)
        {
            this.source = source;
            this.target = target;
            this.listCount = listCount;
            folders = new List<folder>();
        }

        public void AddFolder(string folderName, string serverRelativeUrl, bool isSharePointList, int itemCount)
        {
            folders.Add(new folder(folderName, serverRelativeUrl, isSharePointList, itemCount));
        }

        public void AddFolder(string folderName, string serverRelativeUrl, int itemCount)
        {
            folders.Add(new folder(folderName, serverRelativeUrl, false, itemCount));
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
    public class folder
    {
        [XmlAttribute("folder")]
        public string folderName;
        [XmlAttribute("url")]
        public string serverRelativeUrl;
        [XmlAttribute("list")]
        public bool isSharePointList;
        [XmlAttribute("count")]
        public int itemCount;

        public folder()
        {
            folderName = string.Empty;
            serverRelativeUrl = string.Empty;
            isSharePointList = false;
            itemCount = 0;
        }

        public folder(string folderName, string serverRelativeUrl, bool isSharePointList, int itemCount)
        {
            this.folderName = folderName;
            this.serverRelativeUrl = serverRelativeUrl;
            this.isSharePointList = isSharePointList;
            this.itemCount = itemCount;
        }
    }
}