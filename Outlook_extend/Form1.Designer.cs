namespace Outlook_extend
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.btnAcessEmail = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.button7 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.Location = new System.Drawing.Point(41, 1);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.TabIndex = 13;
            this.monthCalendar1.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateChanged);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(325, 12);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 14;
            this.button2.Text = "Start_time";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(325, 118);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 15;
            this.button3.Text = "End_time";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(406, 15);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(132, 20);
            this.textBox1.TabIndex = 16;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(407, 120);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(131, 20);
            this.textBox2.TabIndex = 17;
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(322, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(159, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "Subject filter(Daily report from ...)";
            // 
            // webBrowser1
            // 
            this.webBrowser1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.webBrowser1.Location = new System.Drawing.Point(2, 246);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(858, 271);
            this.webBrowser1.TabIndex = 21;
            this.webBrowser1.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser1_DocumentCompleted);
            // 
            // button6
            // 
            this.button6.Image = global::Outlook_extend.Properties.Resources.icons8_home_26;
            this.button6.Location = new System.Drawing.Point(768, 4);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(58, 44);
            this.button6.TabIndex = 24;
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button5
            // 
            this.button5.Image = global::Outlook_extend.Properties.Resources.icons8_arrow_right_64;
            this.button5.Location = new System.Drawing.Point(768, 118);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(58, 31);
            this.button5.TabIndex = 23;
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            this.button4.Image = global::Outlook_extend.Properties.Resources.icons8_arrow_left_64;
            this.button4.Location = new System.Drawing.Point(768, 80);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(58, 32);
            this.button4.TabIndex = 22;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click_1);
            // 
            // button1
            // 
            this.button1.Image = global::Outlook_extend.Properties.Resources.icons8_send_email_16;
            this.button1.Location = new System.Drawing.Point(572, 100);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(76, 48);
            this.button1.TabIndex = 9;
            this.button1.UseVisualStyleBackColor = true;
            // 
            // btnAcessEmail
            // 
            this.btnAcessEmail.Image = global::Outlook_extend.Properties.Resources.icons8_find_email_24;
            this.btnAcessEmail.Location = new System.Drawing.Point(572, 4);
            this.btnAcessEmail.Name = "btnAcessEmail";
            this.btnAcessEmail.Size = new System.Drawing.Size(76, 47);
            this.btnAcessEmail.TabIndex = 0;
            this.btnAcessEmail.UseVisualStyleBackColor = true;
            this.btnAcessEmail.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(572, 57);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(169, 20);
            this.textBox3.TabIndex = 25;
            this.textBox3.TextChanged += new System.EventHandler(this.textBox3_TextChanged);
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(324, 79);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(210, 20);
            this.textBox4.TabIndex = 26;
            this.textBox4.TextChanged += new System.EventHandler(this.textBox4_TextChanged);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(2, 171);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(850, 75);
            this.richTextBox1.TabIndex = 27;
            this.richTextBox1.Text = "";
            this.richTextBox1.TextChanged += new System.EventHandler(this.richTextBox1_TextChanged);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(378, 147);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(105, 24);
            this.button7.TabIndex = 28;
            this.button7.Text = "Debug";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(859, 551);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.monthCalendar1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnAcessEmail);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.ShowInTaskbar = false;
            this.Text = "HappyWork, Happy life";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.MonthCalendar monthCalendar1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Button btnAcessEmail;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button button7;
    }
}

