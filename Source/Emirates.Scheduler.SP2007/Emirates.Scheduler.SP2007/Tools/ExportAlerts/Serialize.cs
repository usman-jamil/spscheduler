using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Xml.Serialization;

namespace Emirates.Scheduler.SP2007.Tools.Alerts
{
    [Serializable]
    public class notifications
    {
        [XmlElement("site")]
        public List<site> sites;

        public notifications()
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

        [XmlElement("alert")]
        public List<alert> alerts;

        public site()
        {
        }

        public site(string source, string target)
        {
            this.source = source;
            this.target = target;
            alerts = new List<alert>();
        }

        public void AddAlert(string userName, string listName, string eventType, string alertFrequency, string alertType)
        {
            alerts.Add(new alert(userName, listName, eventType, alertFrequency, alertType, 0));
        }

        public void AddAlert(string userName, string listName, string eventType, string alertFrequency, string alertType, int itemId)
        {
            alerts.Add(new alert(userName, listName, eventType, alertFrequency, alertType, itemId));
        }
    }

    [Serializable]
    public class alert
    {
        [XmlAttribute("user")]
        public string userName;
        [XmlAttribute("list")]
        public string listName;
        [XmlAttribute("event")]
        public string eventType;
        [XmlAttribute("frequency")]
        public string frequency;
        [XmlAttribute("type")]
        public string alertType;
        [XmlAttribute("id")]
        public int itemId;

        public alert()
        {
            userName = string.Empty;
            listName = string.Empty;
            eventType = string.Empty;
            frequency = string.Empty;
            alertType = string.Empty;
            itemId = 0;
        }
        
        public alert(string userName, string listName, string eventType, string frequency, string alertType, int itemId)
        {
            this.userName = userName;
            this.listName = listName;
            this.eventType = eventType;
            this.frequency = frequency;
            this.alertType = alertType;
            this.itemId = itemId;
        }
    }
}