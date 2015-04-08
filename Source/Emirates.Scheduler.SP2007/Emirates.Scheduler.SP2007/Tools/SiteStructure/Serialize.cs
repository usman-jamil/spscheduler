using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Xml.Serialization;

namespace Emirates.Scheduler.SP2007.Tools.Structure
{
    [Serializable]
    public class sitestructure
    {
        [XmlElement("Sites")]
        public List<site> sites;

        public sitestructure()
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
        [XmlAttribute("Source")]
        public string source;
        [XmlAttribute("Target")]
        public string target;

        [XmlElement("Site")]
        public List<web> webs;

        public site()
        {
        }

        public site(string source, string target)
        {
            this.source = source;
            this.target = target;
            webs = new List<web>();
        }

        public void AddAlert(string source, string target)
        {
            webs.Add(new web(source, target));
        }
    }

    [Serializable]
    public class web
    {
        [XmlAttribute("Source")]
        public string source;
        [XmlAttribute("Target")]
        public string target;

        public web()
        {
            source = string.Empty;
            target = string.Empty;
        }

        public web(string source, string target)
        {
            this.source = source;
            this.target = target;
        }
    }
}