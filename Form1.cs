using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace Outlook_Sample
{
    public partial class Form1 : Form
    {
        List<Schedule> scheduleList = new List<Schedule>();

        DateTime meetingTime;     // 会議時間
        DateTime nowTimerTime;    // タイマ現在時刻

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("本日の予定");
            comboBox1.Items.Add("本日から7日間の予定");
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;

            button2.Enabled = false;
            button3.Enabled = false;

            meetingTimeTextBox.ReadOnly = true;
            remainTimeTextBox.ReadOnly = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            textBox1.Text = "";
            scheduleList.Clear();

            Microsoft.Office.Interop.Outlook.Application outlook
              = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace ns = outlook.GetNamespace("MAPI");
            MAPIFolder oFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);


            DateTime dt = DateTime.Today;

            string startDate;
            string endDate;

            if (comboBox1.Text == "本日の予定")
            {
                startDate = dt.ToString("yy/MM/dd");
                endDate = startDate;
            } else if(comboBox1.Text == "本日から7日間の予定")
            {
                startDate = dt.ToString("yy/MM/dd");
                endDate = dt.AddDays(7).ToString("yy/MM/dd");
            } else
            {
                startDate = dt.ToString("yy/MM/dd");
                endDate = startDate;
            }

            string filter = "[Start] >= '" + startDate + "' AND [Start] <= '" + endDate + "'";

            //開始日、終了日の間の予定で絞り込むとき
            //string filter = "[Start] >= '" + startDate + "' AND [Start] <= '" + endDate + "'";
            //予定の題名を「元旦」で絞り込む場合

            //string filter = "[Subject] = '元日'"; 

            //終日の予定で絞り込む場合
            //string filter = "[AllDayEvent] = True";

            Items oItems = oFolder.Items.Restrict(filter);

            textBox1.Text += oFolder.Name + "\r\n";

            AppointmentItem oAppoint = oItems.GetFirst();
            while (oAppoint != null)
            {
                textBox1.Text += oAppoint.Subject + "\r\n";
                textBox1.Text += oAppoint.Start.ToString("yyyy/MM/dd hh:mm:ss") + "\r\n";
                textBox1.Text += oAppoint.End.ToString("yyyy/MM/dd hh:mm:ss") + "\r\n";
                textBox1.Text += "---\r\n";
                Schedule schedule = new Schedule(oAppoint.Subject, oAppoint.Start, oAppoint.End);
                oAppoint = oItems.GetNext();

                scheduleList.Add(schedule);
            }
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void comboBox2_DropDown(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();

            foreach (var list in scheduleList)
            {
                comboBox2.Items.Add(list.subject);
            }
        }


        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string currentComboString;

            textBox2.Clear();

            if (comboBox2.SelectedIndex != -1)
            {
                currentComboString = comboBox2.Items[comboBox2.SelectedIndex].ToString();   // 現在のコンボボックスの文字列を保存

                foreach (var list in scheduleList){
                    if(list.subject == currentComboString)
                    {
                        textBox2.Text += list.subject + "\r\n";
                        textBox2.Text += list.start.ToString("yyyy/MM/dd hh:mm:ss") + "\r\n";
                        textBox2.Text += list.end.ToString("yyyy/MM/dd hh:mm:ss") + "\r\n";

                        if (list.start > System.DateTime.Now)
                        {
                            button2.Enabled = true;
                            meetingTimeTextBox.Text = list.end.ToString("yyyy/MM/dd hh:mm:ss");
                            meetingTime = list.start;
                        }
                    }
                }
            }
        }

        private void timerControl_Tick(object sender, EventArgs e)
        {
            // 経過時間に1秒を加える
            nowTimerTime = nowTimerTime.AddSeconds(1);

            TimeSpan ts = meetingTime - nowTimerTime;                       // 会議開始時間と現在時間の差分を求める

            // 残り時間を表示
            String tempRemain = ts.ToString();
            char[] delimiterChars = { '.', ':' };

            string[] words = tempRemain.Split(delimiterChars);

            int numDays;

            if (!int.TryParse(words[0], out numDays))
            {
                numDays = 0;
            }
            else
            {
                numDays = int.Parse(words[0]);
            }

            words[1] = (int.Parse(words[1]) + numDays * 24).ToString();     // 残り日数は、1日当たり24時間として計算して表示する

            remainTimeTextBox.Text = words[1] + ":" + words[2] + ":" + words[3];

            if (meetingTime < nowTimerTime)
            {
                timerControl.Stop();

                MessageBox.Show("時間になりました");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string MessageBoxTitle = "タイマースタート確認";
            string MessageBoxContent = "タイマーをスタートしてもよろしですか？";

            DialogResult dialogResult = MessageBox.Show(MessageBoxContent, MessageBoxTitle, MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                nowTimerTime = System.DateTime.Now;

                // タイマースタート
                timerControl.Start();

                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = true;
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
            }
            else if (dialogResult == DialogResult.No)
            {

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string MessageBoxTitle = "タイマー解除確認";
            string MessageBoxContent = "タイマーを解除してもよろしですか？";

            DialogResult dialogResult = MessageBox.Show(MessageBoxContent, MessageBoxTitle, MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                // タイマーストップ
                timerControl.Stop();

                button1.Enabled = true;
                button2.Enabled = false;
                button3.Enabled = false;
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;
            }
            else if (dialogResult == DialogResult.No)
            {
                
            }
        }
    }
}

