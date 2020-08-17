using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Diagnostics;


namespace Outlook_extend
{
    
    
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Outlook.Application app = null;
        Microsoft.Office.Interop.Outlook._NameSpace ns = null;
        Microsoft.Office.Interop.Outlook.Items sendEmailItems = null;

        /// <summary>
        /// start data means one filter with email's creation time start point
        /// end data means one filter with email's creation time end point
        /// </summary>
        /// 

        public void authenticationUser(string username, string pwd) { 
        
        
        }


        ///<Unfinished module>
        /// user interface authentication
        /// data cleaning
        ///</Unfinished module>


        public DateTime startDate;
        public DateTime endDate;
        public int endpage = 1;
        //Store email filter by subject
        public SortedDictionary<DateTime, string> tempdic;
        string myPathTemp = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "Results");
        
        // githuub
        public Form1()
        {
            InitializeComponent();
        }

        /*
        create map key:createTime-----value:System.__ComObject
         
        StartTime EndTime Subject filter all send items
         
        */
        public SortedDictionary<DateTime, string> ItemMapInfo(DateTime startDate, DateTime endDate, string subjectString)
        {
            
            
            string tempPath = System.IO.Path.GetTempPath();
            //try delete before files
            try
            {    
            string[] htmlList = Directory.GetFiles(tempPath, "*.html");
            foreach (string f in htmlList)
            {
                File.Delete(f);
            }
            }
            catch (DirectoryNotFoundException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (string.IsNullOrEmpty(subjectString)) {

                return null;
            
            }
            if (string.IsNullOrEmpty(this.startDate.ToString())) {
                return null;
            }

            if (string.IsNullOrEmpty(this.endDate.ToString()))
            {
                return null;
            }
           

            //Store emails filter by start & end time
            //SortedDictionary<DateTime, object> tempdic2 = new SortedDictionary<DateTime, object>();

            app = new Microsoft.Office.Interop.Outlook.Application();
            ns = app.GetNamespace("MAPI");
            ns.Logon("emailadress", "pwd", false, false);

            //Get all sent items list
            sendEmailItems = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail).Items;

            //Count No. of sent items.
            //richTextBox1.AppendText("\n"+"Total send items = " + sendEmailItems.Count);

            try
            {
                
                int countNum = 1;
                
                foreach (var mail in sendEmailItems)
                {
                    Microsoft.Office.Interop.Outlook.MailItem ml = mail as Microsoft.Office.Interop.Outlook.MailItem;
                    //Filter needed  mail by subject in all mails

                    if (ml != null)
                    {
                        if (ml.Subject != null)
                        {
                            if (ml.Subject.ToUpper().Contains(subjectString.ToUpper()))
                            {

                                if (DateTime.Compare(ml.CreationTime, startDate) > 0 && DateTime.Compare(ml.CreationTime, endDate) < 0)
                                {
                                    string mailTempPath = System.IO.Path.Combine(tempPath, countNum + ".html");
                                    
                                    ml.SaveAs(mailTempPath, OlSaveAsType.olHTML);
                                    tempdic.Add(ml.CreationTime, ml.HTMLBody);
                                    
                                    countNum++;
                                }


                            }

                        }

                    }
                    

                }
                if (tempdic.Count() == 0)
                {
                    MessageBox.Show("Please check your selection for subject/time/sendmailbox");
                }

                //Filter needed mail by starttime and endtime in all emails(sorted by time)
                tempdic.OrderBy(KeyValuePair => KeyValuePair.Key);

            }
            catch (ObjectDisposedException ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                ns = null;
                app = null;
            }

            return tempdic;
        }



        private void Form1_Load(object sender, EventArgs e)
        {

        }



        private void button2_Click(object sender, EventArgs e)
        {
            this.startDate = monthCalendar1.SelectionRange.Start;

            textBox1.Text = startDate.ToShortDateString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.endDate = monthCalendar1.SelectionRange.End;
            textBox2.Text = endDate.ToShortDateString();
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Start time : "+textBox1.Text);
            ItemMapInfo(this.startDate, this.endDate, textBox4.Text);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("end time : " + textBox2.Text);
            ItemMapInfo(this.startDate, this.endDate, textBox4.Text);
        }

        private void richTextBox2_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                string tempPath = System.IO.Path.GetTempPath();
                string[] htmlList = Directory.GetFiles(tempPath, "*.html");
                if(this.endpage <= htmlList.Count())
                {
                
                textBox3.AppendText("Path:"+ System.IO.Path.Combine(tempPath, this.endpage + ".html"));


                 webBrowser1.Navigate(System.IO.Path.Combine(tempPath, this.endpage + ".html"));


                this.endpage++;
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                MessageBox.Show(ex.ToString());
            }
           

        }


        // Navigates to the given URL if it is valid.
        private void Navigate(String address)
        {
            if (String.IsNullOrEmpty(address)) return;
            if (address.Equals("about:blank")) return;
            if (!address.StartsWith("http://") &&
                !address.StartsWith("https://"))
            {
                address = "http://" + address;
            }
            try
            {
                webBrowser1.Navigate(new Uri(address));
            }
            catch (System.UriFormatException)
            {
                return;
            }
        }
        //Conllect all email's body to a terminal one
        private void collectionAllHtml(SortedDictionary<DateTime, string> tempdic) 
        {
            
            foreach(var kp in tempdic)
            {

                richTextBox1.AppendText(kp.Value.ToString());

            }
        
        
        }



        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
           
        }
        private void button6_Click(object sender, EventArgs e)
        {
            webBrowser1.GoHome();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            webBrowser1.GoBack();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            webBrowser1.GoForward();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            //CollectionAllEmail();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            collectionAllHtml(this.tempdic);
        }
    }
    

}
