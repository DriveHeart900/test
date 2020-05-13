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

namespace Outlook_extend
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Outlook.Application app = null;
        Microsoft.Office.Interop.Outlook._NameSpace ns = null;
        Microsoft.Office.Interop.Outlook.Items sendEmailItems = null;
        

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            itemMapInfo();

            
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        //create map key:createTime-----value:System.__ComObject


        public SortedDictionary<DateTime, object> itemMapInfo() {

            SortedDictionary<DateTime, object> tempdic = new SortedDictionary<DateTime, object>();

            app = new Microsoft.Office.Interop.Outlook.Application();
            ns = app.GetNamespace("MAPI");
            ns.Logon("emailadress", "pwd", false, false);
            sendEmailItems = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail).Items;


            //sendEmailItems = sendEmailItems.Restrict("CreationTime");
            //sendEmailItems = sendEmailItems.Sort("CreationTime", OlSortOrder.olDescending);
            richTextBox1.AppendText("\n" + sendEmailItems.Count);
            try
            {

                foreach (object mail in sendEmailItems)
                {
                    if ((mail as Microsoft.Office.Interop.Outlook.MailItem) != null)
                    {

                        //richTextBox1.AppendText( );
                        //richTextBox1.AppendText("Subject: " + (mail as Microsoft.Office.Interop.Outlook.MailItem).Subject);
                        //richTextBox1.AppendText("CreationTime: " + (mail as Microsoft.Office.Interop.Outlook.MailItem).CreationTime);
                        //richTextBox1.AppendText("HTMLBody: " + (mail as Microsoft.Office.Interop.Outlook.MailItem).HTMLBody);
                        tempdic.Add((mail as Microsoft.Office.Interop.Outlook.MailItem).CreationTime, (mail as Microsoft.Office.Interop.Outlook.MailItem).Subject + "\n" + (mail as Microsoft.Office.Interop.Outlook.MailItem).Body) ;
                        

                    }
                        
       
                    
                }

                tempdic.OrderBy(KeyValuePair => KeyValuePair.Key);
                foreach (KeyValuePair<DateTime, object> keyValuePair in tempdic)
                {
                    if(DateTime.Compare(keyValuePair.Key, Convert.ToDateTime(dataGridView1.Rows[0].Cells[1].Value)) >0 && DateTime.Compare(keyValuePair.Key, Convert.ToDateTime(dataGridView1.Rows[1].Cells[1].Value))<0)

                    richTextBox1.AppendText("\n" + keyValuePair.Key);
                    richTextBox1.AppendText("\n" + keyValuePair.Value);


                }


                //String path = System.Envrionment.currentDictionery;
                //Path.Combine(Environment.CurrentDirectory, @"topologyInfo.xml"
                //foreach (object mail in tempdic)
                //{
                //    if ((mail as Microsoft.Office.Interop.Outlook.MailItem) != null)
                //    {

                //        (mail as Microsoft.Office.Interop.Outlook.MailItem).Saveas(path,5);


                //    }












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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            
            
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            
           

        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            
        }

       

        private void button2_Click_1(object sender, EventArgs e)
        {

            int MaxRows = 2;

            

            if (dataGridView1.Rows.Count < MaxRows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.AllowUserToAddRows = true;
                dataGridView1.Rows[n].Cells[0].Value = (n + 1).ToString();
                dataGridView1.Rows[n].Cells[1].Value = dateTimePicker1.Value.ToString("MM-dd-yyyy");
                
            }

            
            else
            {
                dataGridView1.AllowUserToAddRows = false;
            }


        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            
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

        public void sortTime() {

        
        }
        

    
    }
}
