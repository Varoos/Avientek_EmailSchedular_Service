using ClosedXML.Excel;
using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;

namespace AVFGService
{
    public partial class Form1 : Form
    {
        GetExternalData clsd = new GetExternalData();
        string subject = "";
        string body = "";
        string Scheduledday = "";
        string scheduledtime = "";
        public Form1()
        {
            InitializeComponent();
        }
        public void GetData()
        {
            try
            {
                System.Data.DataTable dt = clsd.GetToEmailId().Tables[0];
                string smtpaddress = "";
                string FromEmailid = "";
                string pwd = "";
                
                if (dt.Rows.Count > 0)
                {
                    DataSet Fds = clsd.GetFromEmailId();
                    if (Fds.Tables[0].Rows.Count > 0)
                    {
                        smtpaddress = Fds.Tables[0].Rows[0][0].ToString();
                        FromEmailid= Fds.Tables[0].Rows[1][0].ToString();
                        pwd= Fds.Tables[0].Rows[2][0].ToString();
                    }
                    
                    }
                
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["ProcurementEmail"].ToString() != "")
                    {
                        DataSet ds = clsd.GetPendingSalesOrder_Std(Convert.ToInt32(dr["iFilter"]));
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            #region MakeExcel
                            string AppLocation = "";
                            AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                            AppLocation = AppLocation.Replace("file:\\", "");
                            string folderName = AppLocation + "\\ExcelFiles";
                            if (!Directory.Exists(folderName))
                            {
                                Directory.CreateDirectory(folderName);
                            }
                            string file = AppLocation + "\\ExcelFiles\\" + dr["iFilter"].ToString() + "_" + DateTime.Now.ToShortDateString() + ".xlsx";
                            XLWorkbook wb = new XLWorkbook();
                            var ws = wb.Worksheets.Add(dr["iFilter"].ToString());
                            ws.Cell(1, 1).InsertTable(ds.Tables[0]);
                            ws.Tables.FirstOrDefault().ShowAutoFilter = false;
                            ws.Columns().AdjustToContents();
                            wb.SaveAs(file);
                            //using (XLWorkbook wb = new XLWorkbook())
                            //{
                            //    wb.Worksheets.Add(ds.Tables[0]);
                            //    Worksheet ws = wb.Worksheet[0];
                            //    wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            //    wb.Style.Font.Bold = true;
                            //    wb.SaveAs(file);
                            //}


                            #endregion

                            #region MailSection

                            MailMessage mm = new MailMessage();
                            mm.From = new MailAddress(smtpaddress);
                            mm.Subject = subject;
                            mm.Body = body;
                            mm.IsBodyHtml = true;
                            string MultiTo = dr["ProcurementEmail"].ToString();
                            string[] MultiCC = null;
                            string[] MultiBCC = null;

                            mm.To.Add(new MailAddress(MultiTo));




                            mm.Attachments.Add(new Attachment(file));

                            SmtpClient smtp = new SmtpClient();
                            smtp.Host = "smtp.gmail.com";
                            smtp.EnableSsl = true;

                            NetworkCredential NetCred = new NetworkCredential();
                            NetCred.UserName = FromEmailid;
                            NetCred.Password = pwd;
                            smtp.UseDefaultCredentials = true;
                            smtp.Credentials = NetCred;
                            smtp.Port = 587;
                            smtp.Send(mm);
                            #endregion
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                clsd.SetLog(ex.ToString());
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            var currentTime = DateTime.Now;
            try
            {
                DataSet dds = clsd.GetEmailDetails();
                if (dds.Tables[0].Rows.Count > 0)
                {
                    
                    subject = dds.Tables[0].Rows[0]["Subject"].ToString();
                    body = dds.Tables[0].Rows[0]["Body"].ToString();
                    Scheduledday = dds.Tables[0].Rows[0]["ScheduledDay"].ToString();
                    scheduledtime = dds.Tables[0].Rows[0]["ScheduledTime"].ToString();
                    //clsd.SetLog(subject+ body+ Scheduledday+ scheduledtime);
                }
                DateTime t = DateTime.Parse(scheduledtime);
                TimeSpan ts = new TimeSpan();
                ts = t - System.DateTime.Now;
                if (ts.TotalMilliseconds < 0)
                {
                    if (currentTime.DayOfWeek.ToString() == Scheduledday && currentTime.Hour == t.Hour && currentTime.Minute == t.Minute && currentTime.Second == t.Second)
                    {
                        GetData();
                    }
                }
            }
            catch (Exception ex)
            {
                clsd.SetLog(ex.ToString());
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            string intvl = ConfigurationManager.AppSettings["Interval"];
            timer1.Interval = Convert.ToInt32(intvl)*1000;
            timer1.Enabled = true;
            timer1.Tick += new EventHandler(timer1_Tick);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
