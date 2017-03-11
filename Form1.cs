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
                if (oAppoint.IsRecurring)
                {
                    RecurrencePattern pattern = oAppoint.GetRecurrencePattern();

                    // DayOfWeekMask が有効かどうか
                    if (pattern.RecurrenceType == OlRecurrenceType.olRecursWeekly || 
                        pattern.RecurrenceType == OlRecurrenceType.olRecursMonthNth ||
                        pattern.RecurrenceType == OlRecurrenceType.olRecursYearNth) {

                        OlDaysOfWeek mask = pattern.DayOfWeekMask;

                        // どの曜日の周期予定かチェックする
                        if ( (mask & OlDaysOfWeek.olSunday) > 0 )
                        {
                            scheduleList.Add(getRecurrentSchedule(DayOfWeek.Sunday, oAppoint));
                        }

                        if ((mask & OlDaysOfWeek.olMonday) > 0)
                        {
                            scheduleList.Add(getRecurrentSchedule(DayOfWeek.Monday, oAppoint));
                        }

                        if ((mask & OlDaysOfWeek.olTuesday) > 0)
                        {
                            scheduleList.Add(getRecurrentSchedule(DayOfWeek.Tuesday, oAppoint));
                        }

                        if ((mask & OlDaysOfWeek.olWednesday) > 0)
                        {
                            scheduleList.Add(getRecurrentSchedule(DayOfWeek.Wednesday, oAppoint));
                        }

                        if ((mask & OlDaysOfWeek.olThursday) > 0)
                        {
                            scheduleList.Add(getRecurrentSchedule(DayOfWeek.Thursday, oAppoint));
                        }

                        if ((mask & OlDaysOfWeek.olFriday) > 0)
                        {
                            scheduleList.Add(getRecurrentSchedule(DayOfWeek.Friday, oAppoint));
                        }

                        if ((mask & OlDaysOfWeek.olSaturday) > 0)
                        {
                            scheduleList.Add(getRecurrentSchedule(DayOfWeek.Saturday, oAppoint));
                        }
                    }
                }
                else
                {
                    scheduleList.Add(getRecurrentSchedule(oAppoint.Start.DayOfWeek, oAppoint));     // 単発予定の場合は、曜日は、予定開始日時の曜日を格納する
                }

                oAppoint = calendarItemsRestricted.GetNext();
            }
        }

        Schedule getRecurrentSchedule(DayOfWeek targetDayOfWeek, AppointmentItem oAppoint)
        {
            DateTime startAppointment;
            DateTime endAppointment;
            DayOfWeek dayOfWeekAppointment;
            StringBuilder sb = new StringBuilder();

            if (oAppoint.IsRecurring)   // 周期予定であるか
            {
                calculateRecurrentDate(targetDayOfWeek, oAppoint, out startAppointment, out endAppointment, out dayOfWeekAppointment);
                sb.Append("[複] ");
            } else
            {
                startAppointment = oAppoint.Start;
                endAppointment = oAppoint.End;
                dayOfWeekAppointment = targetDayOfWeek;
                sb.Append("[単] ");
            }

            sb.Append(" [" + oAppoint.Subject + "]");
            sb.Append(" [" + startAppointment.ToString("yyyy/MM/dd hh:mm:ss") + "]");
            sb.Append(" [" + endAppointment.ToString("yyyy/MM/dd hh:mm:ss") + "]");
            sb.Append("\r\n");
            textBox1.Text += sb.ToString();
            sb.Clear();

            Schedule schedule = new Schedule(oAppoint.Subject, startAppointment, endAppointment, oAppoint.IsRecurring, dayOfWeekAppointment);
            return schedule;
        }


        void calculateRecurrentDate(DayOfWeek targetDayOfWeek, AppointmentItem oAppoint, out DateTime startRecurrence, out DateTime endRecurrence, out DayOfWeek dayOfWeekRecurrence)
        {
            if (oAppoint.Start < DateTime.Today)  // 周期予定開始日が現在より古かったら
            {
                int diffDayOfWeek = 0;

                // 現在の日時の曜日と周期予定の曜日の差分を求める
                while (true)
                {
                    if(DateTime.Today.AddDays(diffDayOfWeek).DayOfWeek == targetDayOfWeek)
                    {
                        break;
                    }
                    diffDayOfWeek++;
                }

                TimeSpan diffTodayAppointStartDay = DateTime.Today - oAppoint.Start;    // 現在の日時の日にちと繰り返し予定がスタートされる日にちの差分を求める
                int diffDay = diffTodayAppointStartDay.Days + diffDayOfWeek + 1;
                startRecurrence = oAppoint.Start.AddDays(diffDay);
                endRecurrence = oAppoint.End.AddDays(diffDay);
                dayOfWeekRecurrence = targetDayOfWeek;
            }
            else
            {
                startRecurrence = oAppoint.Start;
                endRecurrence = oAppoint.End;
                dayOfWeekRecurrence = targetDayOfWeek;
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

                foreach (var list in scheduleList){

                    if ( (list.start > System.DateTime.Now.Add(diffAleartTime)) && comboBoxItemAppointment.SelectedIndex == listIndex)
                    {
                        btnTimerStart.Enabled = true;

                        string isRecurringState;

                        if (list.isRecurring)
                        {
                            isRecurringState = "[複]";
                        }
                        else
                        {
                            isRecurringState = "[単]";
                        }

                        textBoxSelectedAppointment.Text = 
                            isRecurringState +
                            " [" + list.subject + "]" +
                            " [" + list.start.ToString("yyyy/MM/dd hh:mm:ss") + "]" +
                            " [" + list.end.ToString("yyyy/MM/dd hh:mm:ss") + "]" +
                            " [" + list.dayofWeek.ToString() + "]";

                        meetingTimeTextBox.Text = list.start.ToString("yyyy/MM/dd hh:mm:ss");
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

