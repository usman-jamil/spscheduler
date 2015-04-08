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
        Result iTool.Execute(Job job)
        {
            Result result = new Result(job.Id);

            return result;
        }
    }
}
