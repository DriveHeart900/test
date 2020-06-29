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
        
        String myPath = Environment.CurrentDirectory;
        // githuub
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            ItemMapInfo(textBox1.Text, textBox2.Text, richTextBox2.Text);

            
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }



        /*
        create map key:createTime-----value:System.__ComObject
         
        StartTime EndTime Subject filter all send items
         
        */
        public SortedDictionary<DateTime, object> ItemMapInfo(string StartString, string EndString, string SubjectString) {

            int countNum = 0;

            //Store email filter by subject
            SortedDictionary<DateTime, object> tempdic = new SortedDictionary<DateTime, object>();

            //Store emails filter by start & end time
            //SortedDictionary<DateTime, object> tempdic2 = new SortedDictionary<DateTime, object>();
            
            app = new Microsoft.Office.Interop.Outlook.Application();
            ns = app.GetNamespace("MAPI");
            ns.Logon("emailadress", "pwd", false, false);

            //Get all sent items list
            sendEmailItems = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail).Items;

            //Count No. of sent items.
            richTextBox1.AppendText("\n" + sendEmailItems.Count);
            
            try
            {

                foreach (object mail in sendEmailItems)
                {
                    //Filter needed  mail by subject in all mails
                    
                    if ((mail as Microsoft.Office.Interop.Outlook.MailItem) != null && (mail as Microsoft.Office.Interop.Outlook.MailItem).Subject.Contains(SubjectString) && DateTime.Compare((mail as Microsoft.Office.Interop.Outlook.MailItem).CreationTime, Convert.ToDateTime(textBox1.Text)) > 0 && DateTime.Compare((mail as Microsoft.Office.Interop.Outlook.MailItem).CreationTime, Convert.ToDateTime(textBox2.Text)) < 0)
                    {

                        //richTextBox1.AppendText( );
                        //richTextBox1.AppendText("Subject: " + (mail as Microsoft.Office.Interop.Outlook.MailItem).Subject);
                        //richTextBox1.AppendText("CreationTime: " + (mail as Microsoft.Office.Interop.Outlook.MailItem).CreationTime);
                        //richTextBox1.AppendText("HTMLBody: " + (mail as Microsoft.Office.Interop.Outlook.MailItem).HTMLBody);
                        tempdic.Add((mail as Microsoft.Office.Interop.Outlook.MailItem).CreationTime, (mail as Microsoft.Office.Interop.Outlook.MailItem).Subject + "\n" + (mail as Microsoft.Office.Interop.Outlook.MailItem).HTMLBody);
                        
                        if (tempdic.Count() == 0)
                        {
                            MessageBox.Show("Please check your selection for subject/time/sendmailbox");
                        }
                    }

                  
                    
                }

                //Filter needed mail by starttime and endtime in all emails(sorted by time)
                tempdic.OrderBy(KeyValuePair => KeyValuePair.Key);
                foreach (KeyValuePair<DateTime, object> keyValuePair in tempdic)
                {
                    
                    if(DateTime.Compare(keyValuePair.Key, Convert.ToDateTime(textBox1.Text)) >0 && DateTime.Compare(keyValuePair.Key, Convert.ToDateTime(textBox2.Text)) < 0)
                    {
                        countNum++;
                        richTextBox1.AppendText("\n" + keyValuePair.Key);
                        richTextBox1.AppendText("\n" + keyValuePair.Value);
                    }

                    //tempdic2.Add(keyValuePair.Key, keyValuePair.Value);
                }

                richTextBox1.AppendText("\n all satisfied emails num are " + countNum);


                //foreach (object mail in tempdic2)
                //{
                //    if ((mail as Microsoft.Office.Interop.Outlook.MailItem) != null)
                //    {

                //        (mail as Microsoft.Office.Interop.Outlook.MailItem).Saveas(path,5);


                //}












            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                richTextBox1.AppendText(ex.ToString());
            }
            finally
            {
                ns = null;
                app = null;
            }

            return tempdic;
        }



        //Cumulative method to get list item in weekly report
        /*
            Day01 {issue, progress, teststatus, plan}
            Day01 {issue2, progress2, teststatus2, plan2}
            .
            .
            .
            Weekly {issue + issue2+..., progress+progress2+..., plan+plan2+...}
             */

        public List<string> cumulativefunc(SortedDictionary<string, string> cumuDic) {


            List<string> temp = new List<string>();
            temp = null;
            
            //ensure data add from sorted dictionnary
            /* day01--day02--day03--...*/
            foreach (KeyValuePair<string, string> keyPare in cumuDic) {

                temp.Add(keyPare.Key);
             
                if (!temp.Contains(keyPare.Key)) {

                    temp.Add(keyPare + "---------Resolved");
                    }

                temp.Add(keyPare.Key);
                
                
            
            }


            return temp;


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void button3_Click(object sender, EventArgs e)
        {
            var endDate = monthCalendar1.SelectionRange.End.ToShortDateString();
            textBox2.Text = endDate;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var startDate = monthCalendar1.SelectionRange.Start.ToShortDateString();

            textBox1.Text = startDate;
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.AppendText("Starting to install database...");
        }


        public void InstallSQL()
        {
            try
            {
                Process p = new Process();
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = System.Windows.Forms.Application.StartupPath.Trim() + @"SQL2019-SSEI-Dev.exe";
                //-q[n|b|r|f]   Sets user interface (UI) level:
                //n = no UI
                //b = basic UI (progress only, no prompts)
                //r = reduced UI (dialog at the end of installation)
                //f = full UI
               
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                psi.UseShellExecute = true;
                psi.Verb = "runas";
                psi.Arguments = "/qb username=\"fareast\\v-jili8\" companyname=\"Beyondsoft\" addlocal=ALL  disablenetworkprotocols=\"0\" instancename=\"outlook\" SECURITYMODE=\"SQL\" SAPWD=\"King12#$\"";
                p.StartInfo = psi;
                p.Start();
            }
            catch (System.Exception ee)
            {
                richTextBox1.AppendText(ee.ToString());
            }
        }


    }



    public class myMail {

        public DateTime myDateTime { get; }
        public string mySubject { get; }
        public string myBody { get; }
        public myMail(DateTime mytime, string subject, string body) {
            myDateTime = mytime;
            mySubject = subject;
            myBody = body; 

        }

       


    }
    

}
