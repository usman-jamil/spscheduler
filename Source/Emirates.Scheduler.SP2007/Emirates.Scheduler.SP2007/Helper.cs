using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.SharePoint;

namespace Emirates.Scheduler.SP2007
{
    public sealed class Helper
    {
        static readonly Helper instance = new Helper();

        /* ======================== */

        // Explicit static constructor to tell C# compiler
        // not to mark type as beforefieldinit
        static Helper()
        {
        }

        Helper()
        {
        }

        public static Helper Instance
        {
            get
            {
                return instance;
            }
        }

        public string MapServerRelativeUrl(string url, string source, string target)
        {
            string newServerRelativeUrl = string.Empty;

            source = source.EndsWith("/") ? source.Substring(0, source.Length - 1) : source;
            target = target.EndsWith("/") ? target.Substring(0, source.Length - 1) : target;

            string sourceWebApp = source.Substring(source.IndexOf("/", 7)).ToLower();
            string targetWebApp = target.Substring(target.IndexOf("/", 7)).ToLower();

            newServerRelativeUrl = url.ToLower().Replace(sourceWebApp, targetWebApp);
            return newServerRelativeUrl;
        }
    }
}
