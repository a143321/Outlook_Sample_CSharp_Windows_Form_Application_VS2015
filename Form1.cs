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
        List<AppointmentItem> scheduleList = new List<AppointmentItem>();

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


            // 起動確認デバッグコード //
            //using (TaskService ts = new TaskService())
            //{
            //    // Create a new task
            //    const string taskName = "Test";
            //    Task t = ts.AddTask(taskName,
            //        new TimeTrigger()
            //        {
            //            StartBoundary = System.DateTime.Now.AddMinutes(1),
            //            Enabled = true
            //        },
            //        new ExecAction(directory + notifyExeName, null, null));

            //    // Register the task in the root folder
            //    ts.RootFolder.RegisterTaskDefinition(taskName, t.Definition);
            //}
            // 起動確認デバッグコード//

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

                    sb.Append("[複] ");
                    sb.Append(" [" + oAppoint.Subject + "]");
                    sb.Append(" [" + oAppoint.Start.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                    sb.Append(" [" + oAppoint.End.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                    sb.Append("\r\n");
                    textBox1.Text += sb.ToString();
                    scheduleList.Add(oAppoint);
                }
                else
                {
                    sb.Append("[単] ");
                    sb.Append(" [" + oAppoint.Subject + "]");
                    sb.Append(" [" + oAppoint.Start.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                    sb.Append(" [" + oAppoint.End.ToString("yyyy/MM/dd hh:mm:ss") + "]");
                    sb.Append("\r\n");

                    textBox1.Text += sb.ToString();
                    scheduleList.Add(oAppoint);
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

                if (item.IsRecurring)
                {
                    isRecurringState = "[複] ";
                }
                else
                {
                    isRecurringState = "[単] ";
                }

                comboBoxItemAppointment.Items.Add(isRecurringState + item.Subject + " " + item.Start.ToString("yyyy/MM/dd hh:mm:ss"));
            }
        }

        // 予定コンボボックスのイベントハンドラ
        private void comboBoxItemAppointment_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnTimerStart.Enabled = false;

            if (comboBoxItemAppointment.SelectedIndex != -1)
            {
                int listIndex = 0;

                foreach (var list in scheduleList){

                    if (list.Start > System.DateTime.Now && comboBoxItemAppointment.SelectedIndex == listIndex)
                    {
                        btnTimerStart.Enabled = true;

                        string isRecurringState;

                        if (list.IsRecurring)
                        {
                            isRecurringState = "[複]";
                        }
                        else
                        {
                            isRecurringState = "[単]";
                        }

                        textBoxSelectedAppointment.Text = 
                            isRecurringState +
                            " [" + list.Subject + "]" +
                            " [" + list.Start.ToString("yyyy/MM/dd hh:mm:ss") + "]" +
                            " [" + list.Start.ToString("yyyy/MM/dd hh:mm:ss") + "]"; 

                        meetingTimeTextBox.Text = list.End.ToString("yyyy/MM/dd hh:mm:ss");
                        meetingTime = list.Start;
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

                registerTaskScheduler(alarmTime);
            }
            else if (dialogResult == DialogResult.No)
            {

            }

        }

        private void btnTimerRelease_Click(object sender, EventArgs e)
        {
            string MessageBoxTitle = "タイマー解除確認";
            string MessageBoxContent = "タイマーを解除してもよろしですか？";

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
        private void registerTaskScheduler(DateTime triggerTime)
        {
            using (TaskService ts = new TaskService())
            {
                // Create a new task
                const string taskName = "Test";
                Task t = ts.AddTask(taskName,
                    new TimeTrigger()
                    {
                        StartBoundary = triggerTime,
                        Enabled = true
                    },
                    new ExecAction(directory + notifyExeName, null, null));

                // Register the task in the root folder
                ts.RootFolder.RegisterTaskDefinition(taskName, t.Definition);
            }
        }
    }
}

