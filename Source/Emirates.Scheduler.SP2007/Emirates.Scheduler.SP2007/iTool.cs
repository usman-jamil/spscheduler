using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Emirates.Scheduler.SP2007.Tools;

namespace Emirates.Scheduler.SP2007
{
    public interface iTool
    {
        Result Execute(Job job);
    }
}
