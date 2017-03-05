using System;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32.TaskScheduler;

namespace Outlook_Sample
{
    public partial class Form1 : Form
    {
        List<Schedule> scheduleList = new List<Schedule>();

        DateTime meetingTime;     // 会議開始時刻
        DateTime alarmTime;       // アラーム発生時刻
        DateTime nowTimerTime;    // タイマ現在時刻

        static Assembly myAssembly = Assembly.GetEntryAssembly();
        static string path = myAssembly.Location;
        static string directory = System.IO.Path.GetDirectoryName(path) + "\\";
        static string notifyExeName = "NotifyMeeting.exe";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBoxItemAppointment.DropDownStyle = ComboBoxStyle.DropDownList;

            btnTimerStart.Enabled = false;
            btnTimerRelease.Enabled = false;
            radioButton2.Checked = true;

        }

        private void button1_Click(object sender, EventArgs e)
        {

            textBox1.Text = "";
            scheduleList.Clear();

            Microsoft.Office.Interop.Outlook.Application outlook
              = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace ns = outlook.GetNamespace("MAPI");
            MAPIFolder oFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

            DateTime today = DateTime.Today;
            DateTime endDay = new DateTime(2099, 1, 1, 0, 0, 0);
            // Initial restriction is Jet query for date range
            string filter1 = "[Start] >= '" + today.ToString("g") + "' AND [End] <= '" + endDay.ToString("g") + "'";

            Items calendarItems = oFolder.Items;
            calendarItems.IncludeRecurrences = true;
            Items calendarItemsRestricted = calendarItems.Restrict(filter1);
            calendarItemsRestricted.Sort("[Start]", false);

            AppointmentItem oAppoint = calendarItemsRestricted.GetFirst();

            

            while (oAppoint != null)
            {
                StringBuilder sb = new StringBuilder();

                if (oAppoint.IsRecurring)
                {
                    RecurrencePattern pattern = oAppoint.GetRecurrencePattern();

                    // DayOfWeekMask が有効かどうか
                    if (pattern.RecurrenceType == OlRecurrenceType.olRecursWeekly || 
                        pattern.RecurrenceType == OlRecurrenceType.olRecursMonthNth ||
                        pattern.RecurrenceType == OlRecurrenceType.olRecursYearNth) {

                        // どの曜日の周期予定かチェックする
                        // Sunday    = 00000001(1)
                        // Monday    = 00000010(2)
                        // Tuesday   = 00000100(4)
                        // Wednesday = 00001000(8)
                        // Thursday  = 00010000(16)
                        // Friday    = 00100000(32)
                        // Saturday  = 01000000(64)

                        //public enum DayOfWeek
                        //{
                        //    Friday = 5,
                        //    Monday = 1,
                        //    Saturday = 6,
                        //    Sunday = 0,
                        //    Thursday = 4,
                        //    Tuesday = 2,
                        //    Wednesday = 3
                        //}

                        OlDaysOfWeek mask = pattern.DayOfWeekMask;
                        DateTime startRecurrence;
                        DateTime endRecurrence;
                        DayOfWeek dayOfWeekRecurrence;
                        TimeSpan diffTodayAppointStartDay; // 今日と繰り返し予定スタート差分を求める
                        int diffDayOfWeek;      // 周期予定の曜日と今日の曜日の差分を求める
                        int diffDay;            // 周期予定日と現在日時の差分を求める
                        Schedule schedule;

                        if ( (mask & OlDaysOfWeek.olSunday) > 0 )
                        {
                            if (oAppoint.Start < today)  // 周期予定開始日が現在より古かったら
                            {
                                sb.Append("[単] ");
                                diffDayOfWeek = System.Math.Abs(today.DayOfWeek - DayOfWeek.Sunday);
                                diffTodayAppointStartDay = today - oAppoint.Start;
                                diffDay = diffTodayAppointStartDay.Days + diffDayOfWeek + 1;
                                startRecurrence = oAppoint.Start.AddDays(diffDay);
                                endRecurrence = oAppoint.End.AddDays(diffDay);
                                dayOfWeekRecurrence = DayOfWeek.Sunday;
                            }
                            else
                            {
                                sb.Append("[複] ");
                                startRecurrence = oAppoint.Start;
                                endRecurrence = oAppoint.End;
                                dayOfWeekRecurrence = DayOfWeek.Sunday;
                            }
                            sb.Append("[複] ");
                            sb.Append(" [" + oAppoint.Subject + "]");
                            sb.Append(" [" + startRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append(" [" + endRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append(" [周期曜日 : 日曜日]");
                            sb.Append("\r\n");
                            textBox1.Text += sb.ToString();

                            schedule = new Schedule(oAppoint.Subject, startRecurrence, endRecurrence, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            scheduleList.Add(schedule);
                        }

                        if ((mask & OlDaysOfWeek.olMonday) > 0)
                        {
                            if (oAppoint.Start < today)  // 周期予定開始日が現在より古かったら
                            {
                                sb.Append("[単] ");
                                diffDayOfWeek = System.Math.Abs(today.DayOfWeek - DayOfWeek.Monday);
                                diffTodayAppointStartDay = today - oAppoint.Start;
                                diffDay = diffTodayAppointStartDay.Days + diffDayOfWeek + 1;
                                startRecurrence = oAppoint.Start.AddDays(diffDay);
                                endRecurrence = oAppoint.End.AddDays(diffDay);
                                dayOfWeekRecurrence = DayOfWeek.Monday;
                                schedule = new Schedule(oAppoint.Subject, startRecurrence, endRecurrence, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            }
                            else
                            {
                                sb.Append("[複] ");
                                startRecurrence = oAppoint.Start;
                                endRecurrence = oAppoint.End;
                                dayOfWeekRecurrence = DayOfWeek.Monday;
                                schedule = new Schedule(oAppoint.Subject, oAppoint.Start, oAppoint.End, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            }
                            sb.Append(" [" + oAppoint.Subject + "]");
                            sb.Append(" [" + startRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append(" [" + endRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append("\r\n");
                            textBox1.Text += sb.ToString();
                            schedule = new Schedule(oAppoint.Subject, startRecurrence, endRecurrence, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            scheduleList.Add(schedule);
                        }

                        if ((mask & OlDaysOfWeek.olTuesday) > 0)
                        {
                            if (oAppoint.Start < today)  // 周期予定開始日が現在より古かったら
                            {
                                sb.Append("[単] ");
                                diffDayOfWeek = System.Math.Abs(today.DayOfWeek - DayOfWeek.Tuesday);
                                diffTodayAppointStartDay = today - oAppoint.Start;
                                diffDay = diffTodayAppointStartDay.Days + diffDayOfWeek + 1;
                                startRecurrence = oAppoint.Start.AddDays(diffDay);
                                endRecurrence = oAppoint.End.AddDays(diffDay);
                                dayOfWeekRecurrence = DayOfWeek.Tuesday;
                            }
                            else
                            {
                                sb.Append("[複] ");
                                startRecurrence = oAppoint.Start;
                                endRecurrence = oAppoint.End;
                                dayOfWeekRecurrence = DayOfWeek.Tuesday;
                            }
                            sb.Append(" [" + oAppoint.Subject + "]");
                            sb.Append(" [" + startRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append(" [" + endRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append("\r\n");
                            textBox1.Text += sb.ToString();
                            schedule = new Schedule(oAppoint.Subject, startRecurrence, endRecurrence, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            scheduleList.Add(schedule);
                        }

                        if ((mask & OlDaysOfWeek.olWednesday) > 0)
                        {
                            if (oAppoint.Start < today)  // 周期予定開始日が現在より古かったら
                            {
                                sb.Append("[単] ");
                                diffDayOfWeek = System.Math.Abs(today.DayOfWeek - DayOfWeek.Wednesday);
                                diffTodayAppointStartDay = today - oAppoint.Start;
                                diffDay = diffTodayAppointStartDay.Days + diffDayOfWeek + 1;
                                startRecurrence = oAppoint.Start.AddDays(diffDay);
                                endRecurrence = oAppoint.End.AddDays(diffDay);
                                dayOfWeekRecurrence = DayOfWeek.Wednesday;
                            }
                            else
                            {
                                sb.Append("[複] ");
                                startRecurrence = oAppoint.Start;
                                endRecurrence = oAppoint.End;
                                dayOfWeekRecurrence = DayOfWeek.Wednesday;
                            }
                            sb.Append(" [" + oAppoint.Subject + "]");
                            sb.Append(" [" + startRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append(" [" + endRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append("\r\n");
                            textBox1.Text += sb.ToString();
                            schedule = new Schedule(oAppoint.Subject, startRecurrence, endRecurrence, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            scheduleList.Add(schedule);
                        }

                        if ((mask & OlDaysOfWeek.olThursday) > 0)
                        {
                            if (oAppoint.Start < today)  // 周期予定開始日が現在より古かったら
                            {
                                sb.Append("[単] ");
                                diffDayOfWeek = System.Math.Abs(today.DayOfWeek - DayOfWeek.Thursday);
                                diffTodayAppointStartDay = today - oAppoint.Start;
                                diffDay = diffTodayAppointStartDay.Days + diffDayOfWeek + 1;
                                startRecurrence = oAppoint.Start.AddDays(diffDay);
                                endRecurrence = oAppoint.End.AddDays(diffDay);
                                dayOfWeekRecurrence = DayOfWeek.Thursday;
                            }
                            else
                            {
                                sb.Append("[複] ");
                                startRecurrence = oAppoint.Start;
                                endRecurrence = oAppoint.End;
                                dayOfWeekRecurrence = DayOfWeek.Thursday;
                            }
                            sb.Append(" [" + oAppoint.Subject + "]");
                            sb.Append(" [" + startRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append(" [" + endRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append("\r\n");
                            textBox1.Text += sb.ToString();
                            schedule = new Schedule(oAppoint.Subject, startRecurrence, endRecurrence, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            scheduleList.Add(schedule);
                        }

                        if ((mask & OlDaysOfWeek.olFriday) > 0)
                        {
                            if (oAppoint.Start < today)  // 周期予定開始日が現在より古かったら
                            {
                                sb.Append("[単] ");
                                diffDayOfWeek = System.Math.Abs(today.DayOfWeek - DayOfWeek.Friday);
                                diffTodayAppointStartDay = today - oAppoint.Start;
                                diffDay = diffTodayAppointStartDay.Days + diffDayOfWeek + 1;
                                startRecurrence = oAppoint.Start.AddDays(diffDay);
                                endRecurrence = oAppoint.End.AddDays(diffDay);
                                dayOfWeekRecurrence = DayOfWeek.Friday;
                            }
                            else
                            {
                                sb.Append("[複] ");
                                startRecurrence = oAppoint.Start;
                                endRecurrence = oAppoint.End;
                                dayOfWeekRecurrence = DayOfWeek.Friday;
                            }
                            sb.Append(" [" + oAppoint.Subject + "]");
                            sb.Append(" [" + startRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append(" [" + endRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append("\r\n");
                            textBox1.Text += sb.ToString();
                            schedule = new Schedule(oAppoint.Subject, startRecurrence, endRecurrence, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            scheduleList.Add(schedule);
                        }

                        if ((mask & OlDaysOfWeek.olSaturday) > 0)
                        {

                            if (oAppoint.Start < today)  // 周期予定開始日が現在より古かったら
                            {
                                sb.Append("[単] ");
                                diffDayOfWeek = System.Math.Abs(today.DayOfWeek - DayOfWeek.Saturday);
                                diffTodayAppointStartDay = today - oAppoint.Start;
                                diffDay = diffTodayAppointStartDay.Days + diffDayOfWeek + 1;
                                startRecurrence = oAppoint.Start.AddDays(diffDay);
                                endRecurrence = oAppoint.End.AddDays(diffDay);
                                dayOfWeekRecurrence = DayOfWeek.Saturday;
                                schedule = new Schedule(oAppoint.Subject, startRecurrence, endRecurrence, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            }
                            else
                            {
                                sb.Append("[複] ");
                                startRecurrence = oAppoint.Start;
                                endRecurrence = oAppoint.End;
                                dayOfWeekRecurrence = DayOfWeek.Saturday;
                                schedule = new Schedule(oAppoint.Subject, oAppoint.Start, oAppoint.End, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            }
                            sb.Append(" [" + oAppoint.Subject + "]");
                            sb.Append(" [" + startRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append(" [" + endRecurrence.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                            sb.Append("\r\n");
                            textBox1.Text += sb.ToString();
                            schedule = new Schedule(oAppoint.Subject, startRecurrence, endRecurrence, oAppoint.IsRecurring, dayOfWeekRecurrence);
                            scheduleList.Add(schedule);
                        }
                    }
                }
                else
                {
                    sb.Append("[単] ");
                    sb.Append(" [" + oAppoint.Subject + "]");
                    sb.Append(" [" + oAppoint.Start.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                    sb.Append(" [" + oAppoint.End.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                    sb.Append(" [周期曜日 : 日曜日]");
                    sb.Append("\r\n");

                    textBox1.Text += sb.ToString();
                    Schedule schedule = new Schedule(oAppoint.Subject, oAppoint.Start, oAppoint.End, oAppoint.IsRecurring, oAppoint.Start.DayOfWeek);
                    scheduleList.Add(schedule);
                }

                oAppoint = calendarItemsRestricted.GetNext();
            }
        }

        // 予定コンボボックスのイベントハンドラ
        private void comboBoxItemAppointment_DropDown(object sender, EventArgs e)
        {
            btnTimerStart.Enabled = false;
            comboBoxItemAppointment.Items.Clear();

            foreach (var item in scheduleList)
            {
                string isRecurringState;

                if (item.isRecurring)
                {
                    isRecurringState = "[複] ";
                }
                else
                {
                    isRecurringState = "[単] ";
                }

                comboBoxItemAppointment.Items.Add(isRecurringState + item.subject + " " + item.start.ToString("yyyy/MM/dd hh:mm:ss"));
            }
        }

        // 予定コンボボックスのイベントハンドラ
        private void comboBoxItemAppointment_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnTimerStart.Enabled = false;
            TimeSpan diffAleartTime;

            // 指定したグループ内のラジオボタンでチェックされている物を取り出す
            var RadioButtonChecked_InGroup = groupBox1.Controls.OfType<RadioButton>()
                .SingleOrDefault(rb => rb.Checked == true);

            if (RadioButtonChecked_InGroup.Text == "5分前に通知")
            {
                // アラーム開始時間を会議開始時間の5分前にする
                diffAleartTime = new TimeSpan(0, 5, 0);
            }
            else if (RadioButtonChecked_InGroup.Text == "10分前に通知")
            {
                // アラーム開始時間を会議開始時間の10分前にする
                diffAleartTime = new TimeSpan(0, 10, 0);
            }
            else
            {
                // アラーム開始時間を会議開始時間の15分前にする
                diffAleartTime = new TimeSpan(0, 15, 0);
            }


            if (comboBoxItemAppointment.SelectedIndex != -1)
            {
                int listIndex = 0;
                int infoDayofWeekRecurrence = -1;

                foreach (var list in scheduleList){

                    if ( (list.start > System.DateTime.Now.Add(diffAleartTime)) && comboBoxItemAppointment.SelectedIndex == listIndex)
                    {
                        btnTimerStart.Enabled = true;

                        string isRecurringState;

                        if (list.isRecurring)
                        {
                            isRecurringState = "[複]";
                            infoDayofWeekRecurrence = (int)list.dayofWeekRecurrence;
                        }
                        else
                        {
                            isRecurringState = "[単]";
                        }

                        textBoxSelectedAppointment.Text = 
                            isRecurringState +
                            " [" + list.subject + "]" +
                            " [" + list.start.ToString("yyyy/MM/dd hh:mm:ss") + "]" +
                            " [" + list.end.ToString("yyyy/MM/dd hh:mm:ss") + "]"; 

                        if(infoDayofWeekRecurrence != -1)
                        {
                            textBoxSelectedAppointment.Text += "[" + infoDayofWeekRecurrence.ToString() + "]";
                        }
                        meetingTimeTextBox.Text = list.end.ToString("yyyy/MM/dd hh:mm:ss");
                        meetingTime = list.start;
                    }
                    listIndex++;
                }
            }
        }

        private void timerControl_Tick(object sender, EventArgs e)
        {
            nowTimerTime = nowTimerTime.AddSeconds(1);  // 経過時間に1秒を加える

            TimeSpan ts = alarmTime - nowTimerTime;     // 会議開始時間と現在時間の差分を求める

            String tempRemain = ts.ToString();          // 残り時間を表示

            String hours = (ts.Hours + ts.Days * 24).ToString();    // 残り日数は、1日当たり24時間として計算して表示する

            remainAlarmTimeTextBox.Text = hours + "[時間] " + ts.Minutes + "[分] " + ts.Seconds + "[秒]";

            if (meetingTime < nowTimerTime)
            {
                timerControl.Stop();

                MessageBox.Show("時間になりました");
            }
        }

        private void btnTimerStart_Click(object sender, EventArgs e)
        {
            string MessageBoxTitle = "タイマースタート確認";
            string MessageBoxContent = "タイマーをスタートしてもよろしですか？\r\nなお、Windowsタスクスケジュラーにも同時に登録され、通知時刻に通知されます。";

            DialogResult dialogResult = MessageBox.Show(MessageBoxContent, MessageBoxTitle, MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                nowTimerTime = System.DateTime.Now;

                // 指定したグループ内のラジオボタンでチェックされている物を取り出す
                var RadioButtonChecked_InGroup = groupBox1.Controls.OfType<RadioButton>()
                    .SingleOrDefault(rb => rb.Checked == true);

                if (RadioButtonChecked_InGroup.Text == "5分前に通知")
                {
                    // アラーム開始時間を会議開始時間の5分前にする
                    alarmTime = meetingTime.AddSeconds(-300);
                }
                else if (RadioButtonChecked_InGroup.Text == "10分前に通知")
                {
                    // アラーム開始時間を会議開始時間の10分前にする
                    alarmTime = meetingTime.AddSeconds(-600);
                } else
                {
                    // アラーム開始時間を会議開始時間の15分前にする
                    alarmTime = meetingTime.AddSeconds(-900);
                }
                
                // タイマースタート
                timerControl.Start();

                btnGetSchedule.Enabled = false;
                btnTimerStart.Enabled = false;
                btnTimerRelease.Enabled = true;
                comboBoxItemAppointment.Enabled = false;

                //文字を置換する（「に」を「2」に置換する）
                string taskName = comboBoxItemAppointment.SelectedItem.ToString();
                taskName = taskName.Replace('/', '_');
                taskName = taskName.Replace(':', '_');

                registerTaskScheduler(taskName, alarmTime);
            }
            else if (dialogResult == DialogResult.No)
            {

            }

        }

        private void btnTimerRelease_Click(object sender, EventArgs e)
        {
            string MessageBoxTitle = "タイマー解除確認";
            string MessageBoxContent = "本アプリのタイマーを解除してもよろしですか？\r\nなお、Windowsタスクスケジューラーに登録済みのタスクは解除されません。";

            DialogResult dialogResult = MessageBox.Show(MessageBoxContent, MessageBoxTitle, MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                // タイマーストップ
                timerControl.Stop();

                btnGetSchedule.Enabled = true;
                btnTimerStart.Enabled = false;
                btnTimerRelease.Enabled = false;
                comboBoxItemAppointment.Enabled = true;

                comboBoxItemAppointment_SelectedIndexChanged(sender, e);
            }
            else if (dialogResult == DialogResult.No)
            {
                
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        // Microsoft.Win32.TaskSchedulerのDLLを使用して、タスクスケジューラーに登録する
        private void registerTaskScheduler(string taskName, DateTime triggerTime)
        {
            using (TaskService ts = new TaskService())
            {
                string path = taskName;

                TimeTrigger triger = new TimeTrigger
                {
                    // For scripting, gets or sets the date and time when the trigger is activated.
                    StartBoundary = triggerTime,

                    // For scripting, gets or sets the date and time when the trigger is deactivated. The trigger cannot start the task after it is deactivated.
                    EndBoundary = triggerTime.AddHours(1),

                    // Gets or sets the amount of time that is allowed to complete the task. By default, a task will be stopped 72 hours after it starts to run. You can change this by changing this setting.
                    ExecutionTimeLimit = new TimeSpan(0, 0, 30, 0),

                    Enabled = true
                };

                ExecAction action = new ExecAction(directory + notifyExeName, null, null);

                // Create a new task
                Task t = ts.AddTask(path, triger, action);

                // ITaskSettings::DeleteExpiredTaskAfter
                // 再実行がスケジュールされていない場合に削除されるまでの時間(期間)
                TimeSpan tim = t.Definition.Settings.DeleteExpiredTaskAfter;
                t.Definition.Settings.DeleteExpiredTaskAfter = new TimeSpan(0, 0, 1, 0);

                // 以下のサイトを確認すること
                // http://dynabook.com/assistpc/faq/pcdata/007771.htm
                // Gets or sets a Boolean value that indicates that the task will not be started if the computer is running on battery power.
                t.Definition.Settings.DisallowStartIfOnBatteries = false;

                // Gets or sets a Boolean value that indicates that the task will be stopped if the computer begins to run on battery power.
                t.Definition.Settings.StopIfGoingOnBatteries = false;

                // システムがプロセスに関連付ける優先順位を示します。 
                t.Definition.Settings.Priority = (System.Diagnostics.ProcessPriorityClass)1;

                // Register the task in the root folder
                ts.RootFolder.RegisterTaskDefinition(taskName, t.Definition);
            }
        }
    }
}

