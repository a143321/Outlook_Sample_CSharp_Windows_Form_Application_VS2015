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
        public bool isRecurring;
        public DayOfWeek dayofWeekRecurrence;   // 周期曜日


        // ↓これがコンストラクター
        public Schedule(string subject, DateTime start, DateTime end, bool isRecurring, DayOfWeek dayofWeekRecurrence)
        {
            this.subject = subject;
            this.start = start;
            this.end = end;
            this.isRecurring = isRecurring;
            this.dayofWeekRecurrence = dayofWeekRecurrence;
        }
    }
}