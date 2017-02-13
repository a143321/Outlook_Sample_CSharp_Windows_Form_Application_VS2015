using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace Outlook_Sample
{
    class Schedule
    {
        public string subject;
        public DateTime start;
        public DateTime end;


        // ↓これがコンストラクター
        public Schedule(string subject, DateTime start, DateTime end)
        {
            this.subject = subject;
            this.start = start;
            this.end = end;
        }
    }
}
