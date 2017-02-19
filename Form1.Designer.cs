namespace Outlook_Sample
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnGetSchedule = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.comboBoxItemAppointment = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnTimerStart = new System.Windows.Forms.Button();
            this.timerControl = new System.Windows.Forms.Timer(this.components);
            this.remainAlarmTimeTextBox = new System.Windows.Forms.TextBox();
            this.meetingTimeTextBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnTimerRelease = new System.Windows.Forms.Button();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.textBoxSelectedAppointment = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGetSchedule
            // 
            this.btnGetSchedule.Location = new System.Drawing.Point(12, 13);
            this.btnGetSchedule.Name = "btnGetSchedule";
            this.btnGetSchedule.Size = new System.Drawing.Size(75, 23);
            this.btnGetSchedule.TabIndex = 0;
            this.btnGetSchedule.Text = "予定表取得";
            this.btnGetSchedule.UseVisualStyleBackColor = true;
            this.btnGetSchedule.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 42);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(553, 207);
            this.textBox1.TabIndex = 1;
            // 
            // comboBoxItemAppointment
            // 
            this.comboBoxItemAppointment.FormattingEnabled = true;
            this.comboBoxItemAppointment.Location = new System.Drawing.Point(12, 285);
            this.comboBoxItemAppointment.Name = "comboBoxItemAppointment";
            this.comboBoxItemAppointment.Size = new System.Drawing.Size(395, 20);
            this.comboBoxItemAppointment.TabIndex = 3;
            this.comboBoxItemAppointment.DropDown += new System.EventHandler(this.comboBoxItemAppointment_DropDown);
            this.comboBoxItemAppointment.SelectedIndexChanged += new System.EventHandler(this.comboBoxItemAppointment_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 270);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "選択予定";
            // 
            // btnTimerStart
            // 
            this.btnTimerStart.Location = new System.Drawing.Point(413, 283);
            this.btnTimerStart.Name = "btnTimerStart";
            this.btnTimerStart.Size = new System.Drawing.Size(152, 23);
            this.btnTimerStart.TabIndex = 8;
            this.btnTimerStart.Text = "タイマスタート";
            this.btnTimerStart.UseVisualStyleBackColor = true;
            this.btnTimerStart.Click += new System.EventHandler(this.btnTimerStart_Click);
            // 
            // timerControl
            // 
            this.timerControl.Interval = 1000;
            this.timerControl.Tick += new System.EventHandler(this.timerControl_Tick);
            // 
            // remainAlarmTimeTextBox
            // 
            this.remainAlarmTimeTextBox.Location = new System.Drawing.Point(401, 374);
            this.remainAlarmTimeTextBox.Name = "remainAlarmTimeTextBox";
            this.remainAlarmTimeTextBox.ReadOnly = true;
            this.remainAlarmTimeTextBox.Size = new System.Drawing.Size(164, 19);
            this.remainAlarmTimeTextBox.TabIndex = 9;
            // 
            // meetingTimeTextBox
            // 
            this.meetingTimeTextBox.Location = new System.Drawing.Point(402, 422);
            this.meetingTimeTextBox.Name = "meetingTimeTextBox";
            this.meetingTimeTextBox.ReadOnly = true;
            this.meetingTimeTextBox.Size = new System.Drawing.Size(163, 19);
            this.meetingTimeTextBox.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(400, 359);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(139, 12);
            this.label3.TabIndex = 11;
            this.label3.Text = "アラーム発生までの残り時間";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(405, 407);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 12);
            this.label4.TabIndex = 12;
            this.label4.Text = "会議開始時間";
            // 
            // btnTimerRelease
            // 
            this.btnTimerRelease.Location = new System.Drawing.Point(492, 458);
            this.btnTimerRelease.Name = "btnTimerRelease";
            this.btnTimerRelease.Size = new System.Drawing.Size(73, 23);
            this.btnTimerRelease.TabIndex = 13;
            this.btnTimerRelease.Text = "タイマ解除";
            this.btnTimerRelease.UseVisualStyleBackColor = true;
            this.btnTimerRelease.Click += new System.EventHandler(this.btnTimerRelease_Click);
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(15, 19);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(86, 16);
            this.radioButton1.TabIndex = 14;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "5分前に通知";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton3);
            this.groupBox1.Controls.Add(this.radioButton2);
            this.groupBox1.Controls.Add(this.radioButton1);
            this.groupBox1.Location = new System.Drawing.Point(12, 358);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(362, 46);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "アラーム時刻";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(248, 19);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(92, 16);
            this.radioButton3.TabIndex = 14;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "15分前に通知";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(124, 19);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(92, 16);
            this.radioButton2.TabIndex = 14;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "10分前に通知";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // textBoxSelectedAppointment
            // 
            this.textBoxSelectedAppointment.Location = new System.Drawing.Point(12, 312);
            this.textBoxSelectedAppointment.Name = "textBoxSelectedAppointment";
            this.textBoxSelectedAppointment.ReadOnly = true;
            this.textBoxSelectedAppointment.Size = new System.Drawing.Size(553, 19);
            this.textBoxSelectedAppointment.TabIndex = 16;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(581, 492);
            this.Controls.Add(this.textBoxSelectedAppointment);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnTimerRelease);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.meetingTimeTextBox);
            this.Controls.Add(this.remainAlarmTimeTextBox);
            this.Controls.Add(this.btnTimerStart);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comboBoxItemAppointment);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnGetSchedule);
            this.Name = "Form1";
            this.Text = "予定を忘れないぞう";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGetSchedule;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ComboBox comboBoxItemAppointment;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnTimerStart;
        private System.Windows.Forms.Timer timerControl;
        private System.Windows.Forms.TextBox remainAlarmTimeTextBox;
        private System.Windows.Forms.TextBox meetingTimeTextBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnTimerRelease;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.TextBox textBoxSelectedAppointment;
    }
}

