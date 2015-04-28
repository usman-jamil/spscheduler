using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Xml.Serialization;
using Microsoft.SharePoint;

namespace Emirates.Scheduler.SP2013.Tools
{
    public class ActivateFeatures : iTool
    {
        struct feature
        {
            public string name;
            public Guid definitionId;
            public string scope;

            public feature(string name, Guid definitionId, string scope)
            {
                this.name = name;
                this.definitionId = definitionId;
                this.scope = scope;
            }
        }

        StringBuilder output = null;
        List<feature> siteFeatures = null;
        List<feature> webFeatures = null;
        List<feature> disableSiteFeatures = null;
        List<feature> disableWebFeatures = null;

        public ActivateFeatures()
        {
            output = new StringBuilder();
            siteFeatures = new List<feature>();
            webFeatures = new List<feature>();
            disableSiteFeatures = new List<feature>();
            disableWebFeatures = new List<feature>();
        }

        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);
            string inputXml = job.DownloadAttachment();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputXml);

            XmlNode rootNode = xmlDoc.SelectSingleNode("Sites");

            XmlNodeList siteNodes = xmlDoc.SelectNodes("//Site");
            XmlNodeList featureNodes = xmlDoc.SelectNodes("//Feature");
            XmlNodeList disableFeatureNodes = xmlDoc.SelectNodes("//DisableFeature");

            foreach (XmlNode featureNode in featureNodes)
            {
                string title = featureNode.Attributes["Title"].Value;
                string scope = featureNode.Attributes["Scope"].Value;
                string definitionId = featureNode.Attributes["DefinitionId"].Value;
                if (scope.ToLower().Equals("site"))
                {
                    siteFeatures.Add(new feature(title, new Guid(definitionId), scope));
                }
                else if (scope.ToLower().Equals("web"))
                {
                    webFeatures.Add(new feature(title, new Guid(definitionId), scope));
                }
            }

            foreach (XmlNode disableFeatureNode in disableFeatureNodes)
            {
                string title = disableFeatureNode.Attributes["Title"].Value;
                string scope = disableFeatureNode.Attributes["Scope"].Value;
                string definitionId = disableFeatureNode.Attributes["DefinitionId"].Value;
                if (scope.ToLower().Equals("site"))
                {
                    disableSiteFeatures.Add(new feature(title, new Guid(definitionId), scope));
                }
                else if (scope.ToLower().Equals("web"))
                {
                    disableWebFeatures.Add(new feature(title, new Guid(definitionId), scope));
                }
            }

            foreach (XmlNode siteNode in siteNodes)
            {
                string url = siteNode.Attributes["Target"].Value;

                output.Append(string.Format("updating web: {0}" + Environment.NewLine, url));
                DisableFeatures(url);
                EnableFeatures(url);
            }

            string tmpFile = Scheduler.Instance.CreateTmpFile();

            System.IO.File.WriteAllText(tmpFile, output.ToString());

            result.AddFile(tmpFile);
            return result;
        }

        private void EnableFeatures(string site)
        {
            try
            {
                using (SPWeb webNew = new SPSite(site).OpenWeb())
                {
                    bool isRootWeb = webNew.IsRootWeb;
                    if (isRootWeb)
                    {
                        foreach(feature sfeature in siteFeatures)
                        {
                            try
                            {
                                using (SPWeb web = new SPSite(site).OpenWeb())
                                {
                                    SPFeatureCollection siteFeaturesCollection = web.Site.Features;
                                    output.Append(string.Format("activating site feature: {0}" + Environment.NewLine, sfeature.name));
                                    //var feature = siteFeaturesCollection.SingleOrDefault(f => f.DefinitionId == sfeature.definitionId);
                                    //output.Append(string.Format("feature found: {0}" + Environment.NewLine, feature.Definition.DisplayName));
                                    siteFeaturesCollection.Add(sfeature.definitionId, true);
                                }
                            }
                            catch { output.Append(string.Format("error activating site feature: {0}" + Environment.NewLine, sfeature.name)); }
                        }
                    }

                    foreach (feature wfeature in webFeatures)
                    {
                        try
                        {
                            using (SPWeb web = new SPSite(site).OpenWeb())
                            {
                                SPFeatureCollection webFeaturesCollection = web.Features;
                                //var feature = webFeaturesCollection.SingleOrDefault(f => f.DefinitionId == wfeature.definitionId);
                                output.Append(string.Format("activating web feature: {0}" + Environment.NewLine, wfeature.name));
                                webFeaturesCollection.Add(wfeature.definitionId, true);
                            }
                        }
                        catch { output.Append(string.Format("error activating web feature: {0}" + Environment.NewLine, wfeature.name)); }
                    }
                }
            }
            catch { }
        }

        private void DisableFeatures(string site)
        {
            try
            {
                using (SPWeb webNew = new SPSite(site).OpenWeb())
                {
                    bool isRootWeb = webNew.IsRootWeb;
                    if (isRootWeb)
                    {
                        foreach (feature sfeature in disableSiteFeatures)
                        {
                            try
                            {
                                using (SPWeb web = new SPSite(site).OpenWeb())
                                {
                                    SPFeatureCollection siteFeaturesCollection = web.Site.Features;
                                    //var feature = siteFeaturesCollection.SingleOrDefault(f => f.DefinitionId == sfeature.definitionId);
                                    output.Append(string.Format("de-activating site feature: {0}" + Environment.NewLine, sfeature.name));
                                    siteFeaturesCollection.Remove(sfeature.definitionId, true);
                                }
                            }
                            catch { output.Append(string.Format("error de-activating site feature: {0}" + Environment.NewLine, sfeature.name)); }
                        }
                    }

                    foreach (feature wfeature in disableWebFeatures)
                    {
                        try
                        {
                            using (SPWeb web = new SPSite(site).OpenWeb())
                            {
                                SPFeatureCollection webFeaturesCollection = web.Features;
                                //var feature = webFeaturesCollection.SingleOrDefault(f => f.DefinitionId == wfeature.definitionId);
                                output.Append(string.Format("de-activating web feature: {0}" + Environment.NewLine, wfeature.name));
                                webFeaturesCollection.Remove(wfeature.definitionId, true);
                            }
                        }
                        catch { output.Append(string.Format("error de-activating web feature: {0}" + Environment.NewLine, wfeature.name)); }
                    }
                }
            }
            catch { }
        }
    }
}
