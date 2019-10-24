using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunduszDomowy.Exc
{
    class NoWorkbookExc : Exception
    {
        public NoWorkbookExc() { }
        public NoWorkbookExc(string sWorkbookName) 
            : base (string.Format("Workbook '{0}' doesn't exist!", sWorkbookName))
        {

        }
    }
}
