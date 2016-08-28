using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using slideQ.Model;

namespace slideQ.SmellDetectors
{
    interface ISmellDetector
    {
        List<PresentationSmell> detect();
    }
}
