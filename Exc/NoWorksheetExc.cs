using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunduszDomowy.Exc
{
    class NoWorksheetExc : Exception
    {
        public NoWorksheetExc() { }
        public NoWorksheetExc(string sWorksheetName)
            : base(string.Format("Worksheet '{0}' doesn't exist!", sWorksheetName))
        {

        }
    }
}
