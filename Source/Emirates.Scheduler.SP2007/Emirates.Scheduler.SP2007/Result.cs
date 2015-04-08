using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Emirates.Scheduler.SP2007.Tools;
using System.IO;

namespace Emirates.Scheduler.SP2007
{
    public class Result
    {
        private int id;
        private string[] outputFiles;

        public Result(int id)
        {
            this.id = id;
            this.outputFiles = new string[] { };
        }

        public void AddFile(string fileName)
        {
            List<string> files = new List<string>(outputFiles);
            files.Add(fileName);
            outputFiles = files.ToArray();
        }

        public string[] GetAllFiles()
        {
            return outputFiles;
        }

        public int Id
        {
            get
            {
                return this.id;
            }
        }
    }
}
