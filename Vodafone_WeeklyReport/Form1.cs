using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.Data.SqlClient;
using ClosedXML.Excel;
using System.Text;
using System.IO;
using System.Threading;

namespace Vodafone_WeeklyReport
{
    public partial class Form1 : Form
    {
        StringBuilder myBuilder = new StringBuilder();
        StringBuilder myBuilder1 = new StringBuilder();
        StringBuilder myBuilder2 = new StringBuilder();
        //*************EDI***********
        StringBuilder myBuilderEDI = new StringBuilder();
        StringBuilder myBuilder1EDI = new StringBuilder();
        StringBuilder myBuilder2EDI = new StringBuilder();

        StringBuilder myBuilderSummaryReport = new StringBuilder();

        string conString;
        string SqlQuery;
        string conString1;
        string SqlQuery1;
        DateTime From;
        DateTime To;

        VodafoneDownloadSchedular downloadObj = new VodafoneDownloadSchedular();
        int i = 0;
        public Form1()
        {
            InitializeComponent();
            this.Load += new System.EventHandler(this.Form1_Load);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Start BackgroundWorker
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            TimeSpan now = TimeSpan.Parse(DateTime.Now.TimeOfDay.ToString("hh\\:mm\\:ss"));
            DayOfWeek dayToday = DateTime.Now.DayOfWeek;
            List<DayOfWeek> GlobalRptWeekendDays = new List<DayOfWeek> {
                     DayOfWeek.Saturday, DayOfWeek.Sunday
                              };

            if (GlobalRptWeekendDays.Contains(dayToday))
            {
                //do nothing.
            }
            else
            {
                TimeSpan GlobalRptStart = new TimeSpan(21, 00, 0);// New TimeSpan(6, 45, 0)
                TimeSpan GlobalRptEnd = new TimeSpan(21, 01, 0);

                //TimeSpan GlobalRptStart = new TimeSpan(11, 35, 0);// New TimeSpan(6, 45, 0)
                //TimeSpan GlobalRptEnd = new TimeSpan(11, 36, 0);

                if (now == GlobalRptStart)
                {
                    if (i == 0)
                    {
                        GlobalReport();
                        i++;
                    }
                }
                if (now == GlobalRptEnd)
                {
                    i = 0;
                }

            }
            DayOfWeek dayTodaySummaryReport = DateTime.Now.DayOfWeek;
            List<DayOfWeek> SummaryRptWeekendDays = new List<DayOfWeek> {
                     DayOfWeek.Sunday,DayOfWeek.Monday
                              };

            if (SummaryRptWeekendDays.Contains(dayTodaySummaryReport))
            {
                //do nothing.
            }
            else
            {
                TimeSpan SummaryStart = new TimeSpan(06, 00, 0);// New TimeSpan(6, 45, 0)
                TimeSpan SummaryEnd = new TimeSpan(06, 01, 0);
                if (now == SummaryStart)
                {
                    if (i == 0)
                    {
                        //SummaryReport();
                        i++;
                    }
                }
                if (now == SummaryEnd)
                {
                    i = 0;
                }
            }

            //*************************** Automatic Download**************          

            /*
            if (GlobalRptWeekendDays.Contains(dayToday))
            {
                //do nothing.
            }
            else
            {
                TimeSpan DownloadStart = new TimeSpan(09, 00, 0);// New TimeSpan(6, 45, 0)
                TimeSpan DownloadEnd = new TimeSpan(09, 20, 0);
                DateTime To = DateTime.Now;
                if (now == DownloadStart)
                {
                    if (i == 0)
                    {
                        try
                        {
                           downloadObj.MoveDirectory();
                            downloadObj.copyDirectory();
                            downloadObj.DataProcessing();

                          
                            MailMessage message1 = new MailMessage();
                            SmtpClient smtp = new SmtpClient();
                            //sending mail to rtm support team.
                            //message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                            message1.From = new MailAddress("BLR-OPS-DEV@tangoe.com");


                            message1.To.Add(("namohar.m@tangoe.com"));
                            message1.To.Add(("Madhu.Ramachandra@tangoe.com"));
                           // message1.To.Add(("namohar.m@tangoe.com"));
                            System.Net.Mail.Attachment attachment;
                            message1.Subject = "Vodafone Files were downloaded successfully  for    " + To.ToString("yyyy-MM-dd");

                            StringBuilder sb = new StringBuilder();

                            sb.AppendLine("<h3>Hi Team,<h3>");
                            sb.AppendLine(" ");
                            sb.AppendLine(" ");
                            sb.AppendLine("Vodafone Files were downloaded successfully.");                         
                            message1.Body = sb.ToString();
                            message1.IsBodyHtml = true;
                            smtp.Port = 25;
                            smtp.Host = "outlook-south.tangoe.com";
                            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                            smtp.EnableSsl = false;
                            //smtp.TargetName = "STARTTLS/smtp.office365.com";
                            smtp.Send(message1);
                            i++;
                        }
                        catch (Exception ex)
                        {
                            MailMessage message1 = new MailMessage();
                            SmtpClient smtp = new SmtpClient();
                            //sending mail to rtm support team.
                            //message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                            message1.From = new MailAddress("BLR-OPS-DEV@tangoe.com");

                            message1.To.Add(("namohar.m@tangoe.com"));
                            message1.To.Add(("Madhu.Ramachandra@tangoe.com"));                   

                            message1.Subject = "Vodafone Files were Not downloaded  for    " + To.ToString("yyyy-MM-dd");

                            StringBuilder sb = new StringBuilder();

                            sb.AppendLine("<h3>Hi Team,<h3>");
                            sb.AppendLine(" ");
                            sb.AppendLine(" ");
                            sb.AppendLine("Exception :- "+ex+"");

                            message1.Body = sb.ToString();
                            message1.IsBodyHtml = true;
                            smtp.Port = 25;
                            smtp.Host = "outlook-south.tangoe.com";
                            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                            smtp.EnableSsl = false;
                            // smtp.TargetName = "STARTTLS/smtp.office365.com";
                            smtp.Send(message1);
                            throw ex;
                           
                        }
                    }
                }
                if (now == DownloadEnd)
                {
                    i = 0;
                }
          
            }  */

        }

        private int BackgroundProcessLogicMethod(BackgroundWorker bw, int a)
        {
            int result = 0;
            Thread.Sleep(20000);
            MessageBox.Show("I was doing some work in the background.");
            return result;
        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //if (e.Cancelled) MessageBox.Show("Operation was canceled");
            //else if (e.Error != null) MessageBox.Show(e.Error.Message);
            //else MessageBox.Show(e.Result.ToString());
            // Start BackgroundWorker
            backgroundWorker1.RunWorkerAsync();

        }

        private void GlobalReport()
        {
            int Issue = 0;
            DateTime extraDay = DateTime.Now;
            //delete from table.
            SqlQuery = "delete from REPORT";
            conString = System.Configuration.ConfigurationManager.AppSettings["conString"];
            using (SqlConnection con = new SqlConnection(conString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(SqlQuery, con);
                cmd.ExecuteNonQuery();
            }

            From = Convert.ToDateTime(DateTime.Today.AddDays(-14));
            To = Convert.ToDateTime(DateTime.Now);





            // DataTable pending = GetData("select ReceivedDate, Issue+Invoicesassigned+InvoiceInProgress+invoiceUnassigned+New_Account as PendingTotal, Invoicesassigned+InvoiceInProgress+invoiceUnassigned+New_Account as pending,Issue as Issue from(select ReceivedDate,TotalInvoicesReceived,Invoicesassigned,InvoiceInProgress,(TotalInvoicesReceived - (Invoicesassigned+InvoiceInProgress+QCUnassinged+QCAssigned+QCInprogress+QCCompleted+Duplicate+Issue+EDI+DNP+New_Account+Statement+AlreadyProcessed)) as invoiceUnassigned,(QCUnassinged+QCAssigned+QCInprogress+QCCompleted) as InvoiceCompleted,Duplicate,Issue,EDI,DNP,New_Account,AlreadyProcessed from (select cast(FI_CreatedOn as date )as ReceivedDate,count (FI_OriginalName) as TotalInvoicesReceived,sum([IP_Asg]) as Invoicesassigned,sum([IP_Inp]) as InvoiceInProgress,sum ([QC_Idle]) as QCUnassinged,sum([QC_Asg])as QCAssigned,sum([QC_Inp])as QCInprogress,sum([QC_Comp])as QCCompleted,sum([Duplicate])as Duplicate,sum([IP_Issue])as Issue, SUM([EDI]) as EDI, SUM([DNP]) as DNP, SUM([New_Account]) as New_Account,SUM([Statement]) as Statement,SUM([Already Processed]) as AlreadyProcessed from (select * from (select FI_OriginalName,FI_Source,FI_CreatedOn ,IND_Status,IND_FI  from dbo.Vodafone_FileInfo   left join dbo.Vodafone_InvDetails on IND_FI=FI_ID ) src   pivot (count(IND_Status)  for IND_Status in ([IP_Asg],[QC_Inp],[QC_Idle],[IP_Inp],[QC_Comp],[QC_Asg],[Duplicate],[IP_Issue],[EDI],[DNP],[New_Account],[Statement],[Already Processed]) ) piv) aaa   where FI_Source not in ('EDI') and cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and    cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' group by cast(FI_CreatedOn as date ))a)b ");
            DataTable pending = GetData("select ReceivedDate, Issue+Invoicesassigned+InvoiceInProgress+invoiceUnassigned+New_Account as PendingTotal, Invoicesassigned+InvoiceInProgress+invoiceUnassigned+New_Account as pending,Issue as Issue from(select ReceivedDate,TotalInvoicesReceived,Invoicesassigned,InvoiceInProgress,(TotalInvoicesReceived - (Invoicesassigned+InvoiceInProgress+QCUnassinged+QCAssigned+QCInprogress+QCCompleted+Duplicate+Issue+EDI+DNP+New_Account+Statement+AlreadyProcessed)) as invoiceUnassigned,(QCUnassinged+QCAssigned+QCInprogress+QCCompleted) as InvoiceCompleted,Duplicate,Issue,EDI,DNP,New_Account,AlreadyProcessed from (select cast(FI_CreatedOn as date )as ReceivedDate,count (FI_OriginalName) as TotalInvoicesReceived,sum([IP_Asg]) as Invoicesassigned,sum([IP_Inp]) as InvoiceInProgress,sum ([QC_Idle]) as QCUnassinged,sum([QC_Asg])as QCAssigned,sum([QC_Inp])as QCInprogress,sum([QC_Comp])as QCCompleted,sum([Duplicate])as Duplicate,sum([IP_Issue])as Issue, SUM([EDI]) as EDI, SUM([DNP]) as DNP, SUM([New_Account]) as New_Account,SUM([Statement]) as Statement,SUM([Already Processed]) as AlreadyProcessed from (select * from (select FI_OriginalName,FI_Source,FI_CreatedOn ,IND_Status,IND_FI  from dbo.Vodafone_FileInfo   left join dbo.Vodafone_InvDetails on IND_FI=FI_ID ) src   pivot (count(IND_Status)  for IND_Status in ([IP_Asg],[QC_Inp],[QC_Idle],[IP_Inp],[QC_Comp],[QC_Asg],[Duplicate],[IP_Issue],[EDI],[DNP],[New_Account],[Statement],[Already Processed]) ) piv) aaa   where FI_Source not in ('EDI') and cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and    cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' group by cast(FI_CreatedOn as date ))a)b ");
            DataTable dt = new DataTable();

            //dt = GetData(";with CTE_Input1 as ( select RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate, count(FI_OriginalName) as InvoiceCount from  (select cast(FI_CreatedOn as date ) as RecieveDate, IND_IP_ModifiedOn,FI_OriginalName from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID   where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and FI_Source !='EDI' and IND_Status not in ('IP_Asg','IP_Inp','IP_Issue')) A group by cast(IND_IP_ModifiedOn as date),RecieveDate),CTE_Input2 as (select cast(FI_CreatedOn as date )as RecieveDate, count(FI_OriginalName) as TotalInvoices from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID  where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and FI_Source !='EDI' group by cast(FI_CreatedOn as date )),CTE_Input3 as(SELECT RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,Cast(round(Accuracy,2) as decimal(10,2)) as Day1 FROM (select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input1.InvoiceCount as InvoiceCount,CTE_Input2.TotalInvoices as TotalInvoices, case when CTE_Input1.InvoiceCount<>0 then (CAST(CTE_Input1.InvoiceCount as float)/CAST(CTE_Input2.TotalInvoices as float)*100.0) else 100 END as Accuracy from  CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate )B ) " +
            //             "select * from CTE_Input3");
            dt = GetData(";with CTE_Input1 as ( select RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate, count(FI_OriginalName) as InvoiceCount from  (select cast(FI_CreatedOn as date ) as RecieveDate, IND_IP_ModifiedOn,FI_OriginalName from dbo.Vodafone_FileInfo  left join dbo.Vodafone_InvDetails on IND_FI=FI_ID   where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and FI_Source !='EDI' and IND_Status not in ('IP_Asg','IP_Inp','IP_Issue','New_Account')) A group by cast(IND_IP_ModifiedOn as date),RecieveDate),CTE_Input2 as (select cast(FI_CreatedOn as date )as RecieveDate, count(FI_OriginalName) as TotalInvoices from dbo.Vodafone_FileInfo   left join dbo.Vodafone_InvDetails on IND_FI=FI_ID  where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and FI_Source !='EDI' group by cast(FI_CreatedOn as date )),CTE_Input3 as(SELECT RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,Cast(round(Accuracy,2) as decimal(10,2)) as Day1 FROM (select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input1.InvoiceCount as InvoiceCount,CTE_Input2.TotalInvoices as TotalInvoices, case when CTE_Input1.InvoiceCount<>0 then (CAST(CTE_Input1.InvoiceCount as float)/CAST(CTE_Input2.TotalInvoices as float)*100.0) else 100 END as Accuracy from  CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate )B ) " +
                       "select * from CTE_Input3");
            if (dt.Rows.Count < 1)
            {
                //ScriptManager.RegisterStartupScript(this, GetType(), "YourUniqueScriptKey", "alert('No records found');", true);
                //return;
            }
            //select distinct date.
            var distinctValues = dt.AsEnumerable()
                            .Select(row => new
                            {
                                attribute1_name = row.Field<DateTime>("RecieveDate"),
                            })
                            .Distinct();
            //loop through recieve date
            foreach (var RecieveDate in distinctValues)
            {
                DataTable GroupData = null;

                var query = from t in dt.AsEnumerable()
                            where t.Field<DateTime>("RecieveDate") == RecieveDate.attribute1_name
                            select t;
                if (query != null && query.Count() > 0)
                {
                    int Day6Invoice = 0;
                    decimal Day6Percenatage = 0;
                    int Day5Invoice = 0;
                    decimal Day5Percenatage = 0;
                    int Day4Invoice = 0;
                    decimal Day4Percenatage = 0;
                    int Day3Invoice = 0;
                    decimal Day3Percenatage = 0;
                    int Day2Invoice = 0;
                    decimal Day2Percenatage = 0;
                    int Pending = 0;
                    decimal Day1Percenatage = 0;
                    int Day1Invoice = 0;
                    DateTime Processdate = DateTime.Now;
                    GroupData = query.CopyToDataTable();
                    DateTime Recievedate = Convert.ToDateTime(GroupData.Rows[0]["RecieveDate"]);
                    int TotalInvoices = Convert.ToInt32(GroupData.Rows[0]["TotalInvoices"]);



                    for (int i = 0; i < pending.Rows.Count; i++)
                    {
                        DateTime PendingReceiveDate = Convert.ToDateTime(pending.Rows[i]["ReceivedDate"]);


                        if (PendingReceiveDate == Recievedate)
                        {
                            Pending = Convert.ToInt32(pending.Rows[i]["pending"]);
                            Issue = Convert.ToInt32(pending.Rows[i]["Issue"]);
                        }
                    }

                    int flag = 0;
                    DateTime Day2;
                    DateTime Day3;
                    DateTime Day4;
                    DateTime Day5;
                    DateTime Day6;

                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        flag = 1;
                        Day2 = Recievedate.AddDays(3);
                    }
                    else
                    {
                        Day2 = Recievedate.AddDays(1);
                        flag = 0;
                        if (Day2.DayOfWeek == DayOfWeek.Saturday)
                        {
                            Day2 = Day2.AddDays(2);
                        }
                        if (Day2.DayOfWeek == DayOfWeek.Sunday)
                        {
                            Day2 = Day2.AddDays(1);
                        }
                    }

                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        if (flag == 1)
                        {
                            Day3 = Recievedate.AddDays(4);
                            flag = 1;
                        }
                        else
                        {
                            Day3 = Recievedate.AddDays(3);
                            flag = 0;
                        }
                    }
                    else
                    {
                        if (Day2 == Recievedate.AddDays(3))
                        {
                            Day3 = Recievedate.AddDays(4);
                        }
                        else
                        {
                            Day3 = Recievedate.AddDays(2);
                        }

                        if (Day3.DayOfWeek == DayOfWeek.Saturday)
                        {
                            Day3 = Day3.AddDays(2);
                        }
                        if (Day3.DayOfWeek == DayOfWeek.Sunday)
                        {
                            if (Day2 != Day3.AddDays(1))
                            {
                                Day3 = Day3.AddDays(1);
                            }
                        }
                    }
                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        if (flag == 1)
                        {
                            Day4 = Recievedate.AddDays(5);
                            flag = 1;
                        }
                        else
                        {
                            Day4 = Recievedate.AddDays(3);
                            flag = 0;
                        }
                    }
                    else
                    {
                        if (Day2 == Recievedate.AddDays(3) || Day3 == Recievedate.AddDays(3))
                        {
                            Day4 = Recievedate.AddDays(5);
                        }
                        else
                        {
                            Day4 = Recievedate.AddDays(3);
                        }

                        if (Day4.DayOfWeek == DayOfWeek.Saturday)
                        {
                            Day4 = Day4.AddDays(2);
                        }
                        if (Day4.DayOfWeek == DayOfWeek.Sunday)
                        {
                            if (Day3 != Day4.AddDays(1))
                            {
                                Day4 = Day4.AddDays(1);
                            }
                        }
                    }
                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        if (flag == 1)
                        {
                            Day5 = Recievedate.AddDays(6);
                            flag = 1;
                        }
                        else
                        {
                            Day5 = Recievedate.AddDays(3);
                            flag = 0;
                        }

                    }
                    else
                    {
                        if (Day2 == Recievedate.AddDays(4) || Day3 == Recievedate.AddDays(4) || Day4 == Recievedate.AddDays(4))
                        {
                            Day5 = Recievedate.AddDays(6);
                        }
                        else
                        {
                            Day5 = Recievedate.AddDays(4);
                        }
                        if (Day5.DayOfWeek == DayOfWeek.Saturday)
                        {
                            Day5 = Day5.AddDays(2);
                        }
                        if (Day5.DayOfWeek == DayOfWeek.Sunday)
                        {
                            if (Day4 != Day5.AddDays(1))
                            {
                                Day5 = Day5.AddDays(1);
                            }
                        }


                    }
                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        if (flag == 1)
                        {
                            Day6 = Recievedate.AddDays(7);
                            flag = 1;
                        }
                        else
                        {
                            Day6 = Recievedate.AddDays(3);
                            flag = 0;
                        }

                    }
                    else
                    {
                        Day6 = Recievedate.AddDays(5);
                    }

                    extraDay = Day6;

                    //loop 5 times for day1,day2,day3,day4,day5
                    for (int k = 0; k < GroupData.Rows.Count; k++)
                    {
                        Processdate = Convert.ToDateTime(GroupData.Rows[k]["ProcessDate"]);

                        if (Processdate == Recievedate)
                        {
                            Day1Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day1Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);
                        }

                        //if (GroupData.Rows.Count > 1)
                        if (Processdate == Day2)
                        {
                            Day2Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day2Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);
                        }


                        if (Processdate == Day3)
                        {
                            Day3Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day3Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);

                        }


                        if (Processdate == Day4)
                        {
                            Day4Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day4Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);

                        }


                        if (Processdate == Day5)
                        {
                            Day5Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day5Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);

                        }


                    }



                    SqlQuery = "INSERT INTO REPORT VALUES('" + Recievedate + "'," + TotalInvoices + "," + Pending + "," + Issue + "," + Day1Invoice + "," + Day1Percenatage + "," + Day2Invoice + "," + Day2Percenatage + "," + Day3Invoice + "," + Day3Percenatage + "," + Day4Invoice + "," + Day4Percenatage + "," + Day5Invoice + "," + Day5Percenatage + "," + Day6Invoice + "," + Day6Percenatage + ")";
                    using (SqlConnection con = new SqlConnection(conString))
                    {
                        con.Open();
                        SqlCommand cmd = new SqlCommand(SqlQuery, con);
                        cmd.ExecuteNonQuery();
                    }

                }
            }


            //get 6plus day report
            // DataTable PlusDay_dt = GetData(";with CTE_Input1 as ( select RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate, count(FI_OriginalName) as InvoiceCount from (select cast(FI_CreatedOn as date ) as RecieveDate, IND_IP_ModifiedOn,FI_OriginalName from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID   where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and IND_Status not in ('IP_Asg','IP_Issue')) A group by cast(IND_IP_ModifiedOn as date),RecieveDate),CTE_Input2 as (select cast(FI_CreatedOn as date )as RecieveDate,count(FI_OriginalName) as TotalInvoices from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID  where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' group by cast(FI_CreatedOn as date )),CTE_Input3 as (SELECT RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,Cast(round(Accuracy,2) as decimal(10,2)) as Day1 FROM (select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input1.InvoiceCount as InvoiceCount,CTE_Input2.TotalInvoices as TotalInvoices, case when CTE_Input1.InvoiceCount<>0 then (CAST(CTE_Input1.InvoiceCount as float)/CAST(CTE_Input2.TotalInvoices as float)*100.0) else 100 END as Accuracy from CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate )B ) select  RecieveDate,TotalInvoices,PlusDay as '6PlusDay',Cast(round(Accuracy,2) as decimal(10,2)) as '6PlusDay%' from (select RecieveDate,TotalInvoices,PlusDay,case when PlusDay<>0 then (CAST(PlusDay as float)/CAST(TotalInvoices as float)*100.0) else 100 END as Accuracy  from (select RecieveDate,TotalInvoices,sum(InvoiceCount) as PlusDay from CTE_Input3 WHERE ProcessDate >= DATEADD(day,5, RecieveDate)group by RecieveDate,TotalInvoices)C)D order by RecieveDate");
            DataTable PlusDay_dt = GetData(";with CTE_Input1 as ( select RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate, count(FI_OriginalName) as InvoiceCount from (select cast(FI_CreatedOn as date ) as RecieveDate, IND_IP_ModifiedOn,FI_OriginalName from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID   where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and FI_Source !='EDI' and IND_Status not in ('IP_Asg','IP_Issue')) A group by cast(IND_IP_ModifiedOn as date),RecieveDate),CTE_Input2 as (select cast(FI_CreatedOn as date )as RecieveDate,count(FI_OriginalName) as TotalInvoices from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID  where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and FI_Source !='EDI' group by cast(FI_CreatedOn as date )),CTE_Input3 as (SELECT RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,Cast(round(Accuracy,2) as decimal(10,2)) as Day1 FROM (select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input1.InvoiceCount as InvoiceCount,CTE_Input2.TotalInvoices as TotalInvoices, case when CTE_Input1.InvoiceCount<>0 then (CAST(CTE_Input1.InvoiceCount as float)/CAST(CTE_Input2.TotalInvoices as float)*100.0) else 100 END as Accuracy from CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate )B ) select  RecieveDate,TotalInvoices,PlusDay as '6PlusDay',Cast(round(Accuracy,2) as decimal(10,2)) as '6PlusDay%' from (select RecieveDate,TotalInvoices,PlusDay,case when PlusDay<>0 then (CAST(PlusDay as float)/CAST(TotalInvoices as float)*100.0) else 100 END as Accuracy  from (select RecieveDate,TotalInvoices,sum(InvoiceCount) as PlusDay from CTE_Input3 WHERE ProcessDate >= (select(DATEADD(day, (5 % 5) + CASE ((@@DATEFIRST + DATEPART(weekday, RecieveDate) + (5 % 5)) % 7) WHEN 0 THEN 2 WHEN 1 THEN 1 ELSE 0 END, DATEADD(week, (5 / 5), RecieveDate))) as RecieveDate) group by RecieveDate,TotalInvoices " +
                      " )C)D order by RecieveDate");
            for (int z = 0; z < PlusDay_dt.Rows.Count; z++)
            {
                int plusdayCount = Convert.ToInt32(PlusDay_dt.Rows[z]["6PlusDay"]);
                decimal plusdayPercentage = Convert.ToInt32(PlusDay_dt.Rows[z]["6PlusDay%"]);
                DateTime Recievedate = Convert.ToDateTime(PlusDay_dt.Rows[z]["RecieveDate"]);

                //update 6plus day data.
                SqlQuery = " update Report set Day6Count=" + plusdayCount + ", Day6Percentage='" + plusdayPercentage + "' where RECEIVEDATE='" + Recievedate + "'";
                using (SqlConnection con = new SqlConnection(conString))
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(SqlQuery, con);
                    cmd.ExecuteNonQuery();
                }
            }

            //generate table
            //DataTable weeklyReport_dt = GetData("select convert(varchar, RECEIVEDATE, 105) as DateAdded,Day1Count as Day1,Day2Count AS Day2,Day3Count AS Day3,Day4Count as Day4,Day5Count as Day5,TotalInvoice,Convert(varchar(10),day1Percentage)+'%' AS Day1,Convert(varchar(10),Day2Percentage)+'%' AS Day2,Convert(varchar(10),Day3Percentage)+'%' AS Day3,Convert(varchar(10),Day4Percentage)+'%' AS Day4,Convert(varchar(10),Day5Percentage)+'%' AS Day5 from Report");
            DataTable weeklyReport_dt = GetData("select convert(varchar, RECEIVEDATE, 120) as DateAdded,Day1Count as '1Day',Day2Count AS '2Day',Day3Count AS '3Day',Day4Count as '4Day',Day5Count as '5Day',Day6Count as '6DayPlus',TotalInvoice,Convert(varchar(10),day1Percentage)+'%' AS '1Day ', Convert(varchar(10),Day2Percentage)+'%' AS ' 2Day ', Convert(varchar(10),Day3Percentage)+'%' AS '3Day ', Convert(varchar(10),Day4Percentage)+'%' AS '4Day ',Convert(varchar(10),Day5Percentage)+'%' AS '5Day ',Convert(varchar(10),Day6Percentage)+'%' AS '6DayPlus ' from Report");
            //generate Pending Report.
            DataTable pendingReport_dt = GetData("select convert(varchar, RECEIVEDATE, 120) as DateAdded,Pending,Issue from REPORT where Pending<>0 or Issue <>0");
            //generate SLA report.
            DataTable SLA_dt = GetData("select convert(varchar, RECEIVEDATE, 120) as DateAdded,TotalInvoice,ProcessedInvoices as Met_GS_SLA,Convert(varchar(10),(CAST(round(Percentage,2) as decimal(10,2))))+'%' as GS_SLA from (select RECEIVEDATE,TotalInvoice,sum(Day1Count+Day2Count+Day3Count+Day4Count+Day5Count) as ProcessedInvoices,case when sum(Day1Count+Day2Count+Day3Count+Day4Count+Day5Count)<>0 then (CAST(sum(Day1Count+Day2Count+Day3Count+Day4Count+Day5Count) as float)/CAST(TotalInvoice as float)*100.0) else 100 END as Percentage from dbo.REPORT  group by RECEIVEDATE,TotalInvoice ) as a order by RECEIVEDATE");



            //Set DataTable Name which will be the name of Excel Sheet.
            weeklyReport_dt.TableName = "Vodafone Weekly Report";
            pendingReport_dt.TableName = "Vodafone Pending Report";
            SLA_dt.TableName = "Vodafone SLA Report";

            //***********      EDI Report ************************** 
            EDIReport();
            DataTable EDIweeklyReport_dt = GetDataEDI("select convert(varchar, RECEIVEDATE, 120) as DateAdded,Day1Count as '1Day',Day2Count AS '2Day',Day3Count AS '3Day',Day4Count as '4Day',Day5Count as '5Day',Day6Count as '6DayPlus',TotalInvoice,Convert(varchar(10),day1Percentage)+'%' AS '1Day ', Convert(varchar(10),Day2Percentage)+'%' AS ' 2Day ', Convert(varchar(10),Day3Percentage)+'%' AS '3Day ', Convert(varchar(10),Day4Percentage)+'%' AS '4Day ',Convert(varchar(10),Day5Percentage)+'%' AS '5Day ',Convert(varchar(10),Day6Percentage)+'%' AS '6DayPlus ' from Report");
            //generate Pending Report.
            DataTable EDIpendingReport_dt = GetDataEDI("select convert(varchar, RECEIVEDATE, 120) as DateAdded,Pending,Issue from REPORT where Pending<>0 or Issue <>0");
            //generate SLA report.
            DataTable EDISLA_dt = GetDataEDI("select convert(varchar, RECEIVEDATE, 120) as DateAdded,TotalInvoice,ProcessedInvoices as Met_GS_SLA,Convert(varchar(10),(CAST(round(Percentage,2) as decimal(10,2))))+'%' as GS_SLA from (select RECEIVEDATE,TotalInvoice,sum(Day1Count+Day2Count+Day3Count+Day4Count+Day5Count) as ProcessedInvoices,case when sum(Day1Count+Day2Count+Day3Count+Day4Count+Day5Count)<>0 then (CAST(sum(Day1Count+Day2Count+Day3Count+Day4Count+Day5Count) as float)/CAST(TotalInvoice as float)*100.0) else 100 END as Percentage from dbo.REPORT  group by RECEIVEDATE,TotalInvoice ) as a order by RECEIVEDATE");



            //Set DataTable Name which will be the name of Excel Sheet.
            EDIweeklyReport_dt.TableName = "Vodafone-EDI Weekly Report";
            EDIpendingReport_dt.TableName = "Vodafone-EDI Pending Report";
            EDISLA_dt.TableName = "Vodafone-EDI SLA Report";

            //Create a New Workbook.
            using (XLWorkbook wb = new XLWorkbook())
            {
                //Add the DataTable as Excel Worksheet.----Manual
                wb.Worksheets.Add(weeklyReport_dt);
                wb.Worksheets.Add(pendingReport_dt);
                wb.Worksheets.Add(SLA_dt);

                //Add the DataTable as Excel Worksheet.-----EDI
                wb.Worksheets.Add(EDIweeklyReport_dt);
                wb.Worksheets.Add(EDIpendingReport_dt);
                wb.Worksheets.Add(EDISLA_dt);

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    //Save the Excel Workbook to MemoryStream.
                    wb.SaveAs(memoryStream);

                    //Convert MemoryStream to Byte array.
                    byte[] bytes = memoryStream.ToArray();
                    memoryStream.Close();


                    MailMessage message1 = new MailMessage();
                    SmtpClient smtp = new SmtpClient();
                    //sending mail to rtm support team.
                    //message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.From = new MailAddress("BLR-OPS-DEV@tangoe.com");

                    //message1.To.Add(("Manjunath.Jayaram@tangoe.com"));
                    message1.To.Add(("Sriram.Krishnan@tangoe.com"));
                    message1.To.Add(("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(("Johwessly.Chennaiah@tangoe.com"));
                    message1.To.Add(("Madhu.Ramachandra@tangoe.com"));
                    message1.To.Add(("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(("Kyle.Borner@tangoe.com"));
                    message1.To.Add(("nicole.goodwin@tangoe.com"));
                    message1.To.Add(("melissa.guarracino@tangoe.com"));
                    message1.To.Add(("Sherrill.Rieken@tangoe.com"));



                    message1.CC.Add(("Sleema.Joseph@tangoe.com"));
                    message1.CC.Add(("namohar.m@tangoe.com"));
                    // message1.To.Add(("namohar.m@tangoe.com"));

                    System.Net.Mail.Attachment attachment;

                    message1.Attachments.Add(new Attachment(new MemoryStream(bytes), "GlobalSupport Vodafone Invoice Processing EOB Report.xlsx"));


                    message1.Subject = "Global Support Vodafone Invoice Processing EOB Report for    " + To.ToString("yyyy-MM-dd");
                    //-----Manual email body-----
                    getDelayHTML1(pendingReport_dt);
                    getDelayHTML(weeklyReport_dt);
                    getDelayHTML2(SLA_dt);
                    //-----EDI email body-----
                    getDelayHTML1EDI(EDIpendingReport_dt);
                    getDelayHTMLEDI(EDIweeklyReport_dt);
                    getDelayHTML2EDI(EDISLA_dt);
                    StringBuilder sb = new StringBuilder();

                    sb.AppendLine("<h3>Hi Team,<h3>");

                    //sb.AppendLine("");
                    sb.AppendLine(string.Format("<br /> Global Support Vodafone Manual Invoice Processing EOB Report  for : " + To.ToString("yyyy-MM-dd")));
                    sb.AppendLine("<br />");
                    sb.AppendLine("<br />");
                    sb.AppendLine("Pending/In Queue Volumes:");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder1.ToString());
                    sb.AppendLine("");
                    sb.AppendLine("Weekly Report : ");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");
                    sb.AppendLine("SLA Report : ");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder2.ToString());
                    sb.AppendLine("");//here I want the data to display in table formats

                    //sb.AppendLine("Please note: " + "http://10.80.20.84/PendingRecordsEMSDB/Gridview/grid" + " please check this link for pending volumes by client and batch date. ");
                    sb.AppendLine("");
                    sb.AppendLine("<br />");
                    //  ************EDI * ***************************
                    sb.AppendLine("<br />");
                    sb.AppendLine("");
                    sb.AppendLine("********************* Global Support Vodafone Electronic Invoice Processing EOB Report. ****************************");
                    sb.AppendLine("<br />");
                    sb.AppendLine("<br />");
                    sb.AppendLine("<br />");
                    sb.AppendLine("Pending/In Queue Volumes-EDI:");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder1EDI.ToString());
                    sb.AppendLine("");
                    sb.AppendLine("Weekly Report-EDI: ");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilderEDI.ToString());
                    sb.AppendLine("");
                    sb.AppendLine("SLA Report-EDI: ");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder2EDI.ToString());
                    sb.AppendLine("");
                    sb.AppendLine("");
                    //sb.AppendLine(string.Format("<b>Please note: </b>" + "http://10.80.20.84/PendingRecordsEMSDB/Gridview/EDI" + " <b>please check this link for pending volumes by client and batch date. "));
                    sb.AppendLine("");
                    sb.AppendLine("<br />");
                    sb.AppendLine("");
                    sb.AppendLine(string.Format("Please review and let us know if you have any clarifications.<br /><br /><P> Note : This is system generated mail. please send mail to namohar.m@tangoe.com and tangoe-devops@tangoe.com if you have any issues.</p><br /><br /> Thank You.<br /> Namohar <br /> "));
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                    message1.IsBodyHtml = true;
                    smtp.Port = 25;
                    smtp.Host = "outlook-south.tangoe.com";
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.EnableSsl = false;
                    // smtp.TargetName = "STARTTLS/smtp.office365.com";
                    smtp.Send(message1);

                }
            }
        }

        private DataTable GetData(string query)
        {
            //string conString = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;
            string conString = System.Configuration.ConfigurationManager.AppSettings["conString"];

            using (SqlConnection con = new SqlConnection(conString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(query, con);
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    return dt;
                }
            }
        }

        private string getDelayHTML(DataTable dt)
        {

            myBuilder = new StringBuilder();

            myBuilder.Append("<table border='1' cellpadding='5' cellspacing='0' ");
            myBuilder.Append("style='border: solid 1px Silver; font-size: x-small;'>");

            myBuilder.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in dt.Columns)
            {
                myBuilder.Append("<td align='left' valign='top' bgcolor='#3F88F6'>");
                myBuilder.Append("<B />" + myColumn.ColumnName);
                myBuilder.Append("</td>");
            }
            myBuilder.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                if (myRow.ItemArray.Contains("Total"))
                {
                    myBuilder.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                        myBuilder.Append("<B />" + myRow[myColumn.ColumnName].ToString());
                        myBuilder.Append("</td>");
                    }
                    myBuilder.Append("</tr>");
                }
                else
                {
                    myBuilder.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder.Append("<td align='left' valign='top'>");
                        myBuilder.Append(myRow[myColumn.ColumnName].ToString());
                        myBuilder.Append("</td>");
                    }
                    myBuilder.Append("</tr>");
                }

            }
            myBuilder.Append("</table>");

            return myBuilder.ToString();
        }

        private string getDelayHTML1(DataTable dt)
        {

            myBuilder1 = new StringBuilder();

            myBuilder1.Append("<table border='1' cellpadding='5' cellspacing='0' ");
            myBuilder1.Append("style='border: solid 1px Silver; font-size: x-small;'>");

            myBuilder1.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in dt.Columns)
            {
                myBuilder1.Append("<td align='left' valign='top' bgcolor='#3F88F6'>");
                myBuilder1.Append("<B />" + myColumn.ColumnName);
                myBuilder1.Append("</td>");
            }
            myBuilder1.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                if (myRow.ItemArray.Contains("Total"))
                {
                    myBuilder1.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder1.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                        myBuilder1.Append("<B />" + myRow[myColumn.ColumnName].ToString());
                        myBuilder1.Append("</td>");
                    }
                    myBuilder1.Append("</tr>");
                }
                else
                {
                    myBuilder1.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder1.Append("<td align='left' valign='top'>");
                        myBuilder1.Append(myRow[myColumn.ColumnName].ToString());
                        myBuilder1.Append("</td>");
                    }
                    myBuilder1.Append("</tr>");
                }

            }
            myBuilder1.Append("</table>");

            return myBuilder1.ToString();
        }

        private string getDelayHTML2(DataTable dt)
        {

            myBuilder2 = new StringBuilder();

            myBuilder2.Append("<table border='1' cellpadding='5' cellspacing='0' ");
            myBuilder2.Append("style='border: solid 1px Silver; font-size: x-small;'>");

            myBuilder2.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in dt.Columns)
            {
                myBuilder2.Append("<td align='left' valign='top' bgcolor='#3F88F6'>");
                myBuilder2.Append("<B />" + myColumn.ColumnName);
                myBuilder2.Append("</td>");
            }
            myBuilder2.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                if (myRow.ItemArray.Contains("Total"))
                {
                    myBuilder2.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder2.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                        myBuilder2.Append("<B />" + myRow[myColumn.ColumnName].ToString());
                        myBuilder1.Append("</td>");
                    }
                    myBuilder2.Append("</tr>");
                }
                else
                {
                    myBuilder2.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder2.Append("<td align='left' valign='top'>");
                        myBuilder2.Append(myRow[myColumn.ColumnName].ToString());
                        myBuilder2.Append("</td>");
                    }
                    myBuilder2.Append("</tr>");
                }

            }
            myBuilder2.Append("</table>");

            return myBuilder2.ToString();
        }
        //********************EDI*************************

        private DataTable GetDataEDI(string query)
        {
            string conString = System.Configuration.ConfigurationManager.AppSettings["conString"];
            //  string conString = System.Configuration.ConfigurationManager.AppSettings["conStringEDI"];

            using (SqlConnection con = new SqlConnection(conString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(query, con);
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    return dt;
                }
            }
        }
        private void EDIReport()
        {
            int Issue = 0;
            DateTime extraDay = DateTime.Now;
            //delete from table.
            SqlQuery1 = "delete from REPORT";
            //conString1 = System.Configuration.ConfigurationManager.AppSettings["conStringEDI"];
            conString1 = System.Configuration.ConfigurationManager.AppSettings["conString"];
            using (SqlConnection con = new SqlConnection(conString1))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(SqlQuery1, con);
                cmd.ExecuteNonQuery();
            }

            From = Convert.ToDateTime(DateTime.Today.AddDays(-14));
            To = Convert.ToDateTime(DateTime.Now);


            // DataTable pending = GetDataEDI("select ReceivedDate, Issue+Invoicesassigned+InvoiceInProgress+invoiceUnassigned as PendingTotal, Invoicesassigned+InvoiceInProgress+invoiceUnassigned as pending,Issue as Issue from(select ReceivedDate,TotalInvoicesReceived,Invoicesassigned,InvoiceInProgress,(TotalInvoicesReceived - (Invoicesassigned+InvoiceInProgress+QCUnassinged+QCAssigned+QCInprogress+QCCompleted+Duplicate+Issue+EDI+DNP+Expedite+Statement)) as invoiceUnassigned,(QCUnassinged+QCAssigned+QCInprogress+QCCompleted) as InvoiceCompleted,Duplicate,Issue,EDI,DNP,Expedite from (select cast(FI_CreatedOn as date )as ReceivedDate,count (FI_OriginalName) as TotalInvoicesReceived,sum([IP_Asg]) as Invoicesassigned,sum([IP_Inp]) as InvoiceInProgress,sum ([QC_Idle]) as QCUnassinged,sum([QC_Asg])as QCAssigned,sum([QC_Inp])as QCInprogress,sum([QC_Comp])as QCCompleted,sum([Duplicate])as Duplicate,sum([IP_Issue])as Issue, SUM([EDI]) as EDI, SUM([DNP]) as DNP, SUM([Expedite]) as Expedite,SUM([Statement]) as Statement from (select * from (select FI_OriginalName,FI_CreatedOn ,IND_Status,IND_FI  from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID ) src   pivot (count(IND_Status)  for IND_Status in ([IP_Asg],[QC_Inp],[QC_Idle],[IP_Inp],[QC_Comp],[QC_Asg],[Duplicate],[IP_Issue],[EDI],[DNP],[Expedite],[Statement]) ) piv) aaa   where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and    cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' group by cast(FI_CreatedOn as date ))a)b ");
            DataTable pending = GetData("select ReceivedDate,Issue+Invoicesassigned+InvoiceInProgress as PendingTotal, Invoicesassigned+InvoiceInProgress as pending,Issue as Issue from (select cast(FI_CreatedOn as date )as ReceivedDate,sum(cast(IND_InvoiceNo as int))  as TotalInvoicesReceived,sum([IP_Asg]) as Invoicesassigned,sum([IP_Inp]) as InvoiceInProgress,sum ([QC_Idle]) as QCUnassinged,sum([QC_Asg])as QCAssigned,sum([QC_Inp])as QCInprogress,sum([QC_Comp])as QCCompleted,sum([Duplicate])as Duplicate,sum([IP_Issue])as Issue, SUM([EDI]) as EDI, SUM([DNP]) as DNP, SUM([Expedite]) as Expedite,SUM([Statement]) as Statement from (select * from (select FI_OriginalName,FI_CreatedOn ,IND_Status,IND_FI,IND_InvoiceNo  from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID where FI_Source in ('EDI')) src pivot (count(IND_Status)  for IND_Status in ([IP_Asg],[QC_Inp],[QC_Idle],[IP_Inp],[QC_Comp],[QC_Asg],[Duplicate],[IP_Issue],[EDI],[DNP],[Expedite],[Statement]) ) piv) aaa  where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "'  group by cast(FI_CreatedOn as date ))B");
            DataTable dt = new DataTable();


            //  dt = GetDataEDI(";with CTE_Input1 as (select cast(FI_CreatedOn as date ) as RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate,IND_InvoiceNo  from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_Status not in ('IP_Asg','IP_Inp','IP_Issue')and IND_IP_Processed_By !='null' and FI_Source in ('EDI') ),CTE_Input2 as (select cast(FI_CreatedOn as date ) as RecieveDate, sum(cast(IND_InvoiceNo as int)) as TotalInvoices  from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and FI_Source in ('EDI') group by cast(FI_CreatedOn as date )),CTE_Input3 as(select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input2.TotalInvoices as TotalInvoices,sum(cast(CTE_Input1.IND_InvoiceNo as int)) as InvoiceCount  from CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate  group by CTE_Input1.RecieveDate,CTE_Input1.ProcessDate,TotalInvoices),CTE_Input4 as(select RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,case when InvoiceCount<>0 then (CAST(InvoiceCount as float)/CAST(TotalInvoices as float)*100.0) else 100 END as Accuracy from CTE_Input3 ), CTE_Input5 as (select RecieveDate,sum(InvoiceCount)as Pending from CTE_Input4 group by RecieveDate),CTE_Input6 as( select  CTE_Input4.RecieveDate as RecieveDate,ProcessDate,TotalInvoices,InvoiceCount, Cast(round(Accuracy,2) as decimal(10,2)) as Day1,CTE_Input5.Pending as Pending from CTE_Input5 join CTE_Input4 on CTE_Input4.RecieveDate=CTE_Input5.RecieveDate) select  RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,Day1,case when TotalInvoices<>0 then (CAST(TotalInvoices as int)-CAST(Pending as int)) else 0 END as Pending  from CTE_Input6");
            dt = GetData(";with CTE_Input1 as (select cast(FI_CreatedOn as date ) as RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate,IND_InvoiceNo  from dbo.Vodafone_FileInfo  left join dbo.Vodafone_InvDetails on IND_FI=FI_ID where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_Status not in ('IP_Asg','IP_Inp','IP_Issue')and IND_IP_Processed_By !='null' and FI_Source in ('EDI') ),CTE_Input2 as (select cast(FI_CreatedOn as date ) as RecieveDate, sum(cast(IND_InvoiceNo as int)) as TotalInvoices  from dbo.Vodafone_FileInfo  left join dbo.Vodafone_InvDetails on IND_FI=FI_ID where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and FI_Source in ('EDI') group by cast(FI_CreatedOn as date )),CTE_Input3 as(select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input2.TotalInvoices as TotalInvoices,sum(cast(CTE_Input1.IND_InvoiceNo as int)) as InvoiceCount  from CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate  group by CTE_Input1.RecieveDate,CTE_Input1.ProcessDate,TotalInvoices),CTE_Input4 as(select RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,case when InvoiceCount<>0 then (CAST(InvoiceCount as float)/CAST(TotalInvoices as float)*100.0) else 100 END as Accuracy from CTE_Input3 ), CTE_Input5 as (select RecieveDate,sum(InvoiceCount)as Pending from CTE_Input4 group by RecieveDate),CTE_Input6 as( select  CTE_Input4.RecieveDate as RecieveDate,ProcessDate,TotalInvoices,InvoiceCount, Cast(round(Accuracy,2) as decimal(10,2)) as Day1,CTE_Input5.Pending as Pending from CTE_Input5 join CTE_Input4 on CTE_Input4.RecieveDate=CTE_Input5.RecieveDate) select  RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,Day1,case when TotalInvoices<>0 then (CAST(TotalInvoices as int)-CAST(Pending as int)) else 0 END as Pending  from CTE_Input6");
            if (dt.Rows.Count < 1)
            {
                //ScriptManager.RegisterStartupScript(this, GetType(), "YourUniqueScriptKey", "alert('No records found');", true);
                //return;
            }
            //select distinct date.
            var distinctValues = dt.AsEnumerable()
                            .Select(row => new
                            {
                                attribute1_name = row.Field<DateTime>("RecieveDate"),
                            })
                            .Distinct();
            //loop through recieve date
            foreach (var RecieveDate in distinctValues)
            {
                DataTable GroupData = null;

                var query = from t in dt.AsEnumerable()
                            where t.Field<DateTime>("RecieveDate") == RecieveDate.attribute1_name
                            select t;
                if (query != null && query.Count() > 0)
                {
                    int Day6Invoice = 0;
                    decimal Day6Percenatage = 0;
                    int Day5Invoice = 0;
                    decimal Day5Percenatage = 0;
                    int Day4Invoice = 0;
                    decimal Day4Percenatage = 0;
                    int Day3Invoice = 0;
                    decimal Day3Percenatage = 0;
                    int Day2Invoice = 0;
                    decimal Day2Percenatage = 0;
                    int Pending = 0;
                    decimal Day1Percenatage = 0;
                    int Day1Invoice = 0;
                    DateTime Processdate = DateTime.Now;
                    GroupData = query.CopyToDataTable();
                    DateTime Recievedate = Convert.ToDateTime(GroupData.Rows[0]["RecieveDate"]);
                    int TotalInvoices = Convert.ToInt32(GroupData.Rows[0]["TotalInvoices"]);

                    for (int i = 0; i < pending.Rows.Count; i++)
                    {
                        DateTime PendingReceiveDate = Convert.ToDateTime(pending.Rows[i]["ReceivedDate"]);

                        if (PendingReceiveDate == Recievedate)
                        {
                            Pending = Convert.ToInt32(pending.Rows[i]["pending"]);
                            Issue = Convert.ToInt32(pending.Rows[i]["Issue"]);
                        }
                    }


                    int flag = 0;
                    DateTime Day2;
                    DateTime Day3;
                    DateTime Day4;
                    DateTime Day5;
                    DateTime Day6;


                    //extraDay = Day6;
                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        flag = 1;
                        Day2 = Recievedate.AddDays(3);
                    }
                    else
                    {
                        Day2 = Recievedate.AddDays(1);
                        flag = 0;
                        if (Day2.DayOfWeek == DayOfWeek.Saturday)
                        {
                            Day2 = Day2.AddDays(2);
                        }
                        if (Day2.DayOfWeek == DayOfWeek.Sunday)
                        {
                            Day2 = Day2.AddDays(1);
                        }
                    }

                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        if (flag == 1)
                        {
                            Day3 = Recievedate.AddDays(4);
                            flag = 1;
                        }
                        else
                        {
                            Day3 = Recievedate.AddDays(3);
                            flag = 0;
                        }
                    }
                    else
                    {
                        if (Day2 == Recievedate.AddDays(3))
                        {
                            Day3 = Recievedate.AddDays(4);
                        }
                        else
                        {
                            Day3 = Recievedate.AddDays(2);
                        }

                        if (Day3.DayOfWeek == DayOfWeek.Saturday)
                        {
                            Day3 = Day3.AddDays(2);
                        }
                        if (Day3.DayOfWeek == DayOfWeek.Sunday)
                        {
                            if (Day2 != Day3.AddDays(1))
                            {
                                Day3 = Day3.AddDays(1);
                            }
                        }
                    }
                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        if (flag == 1)
                        {
                            Day4 = Recievedate.AddDays(5);
                            flag = 1;
                        }
                        else
                        {
                            Day4 = Recievedate.AddDays(3);
                            flag = 0;
                        }
                    }
                    else
                    {
                        if (Day2 == Recievedate.AddDays(3) || Day3 == Recievedate.AddDays(3))
                        {
                            Day4 = Recievedate.AddDays(5);
                        }
                        else
                        {
                            Day4 = Recievedate.AddDays(3);
                        }

                        if (Day4.DayOfWeek == DayOfWeek.Saturday)
                        {
                            Day4 = Day4.AddDays(2);
                        }
                        if (Day4.DayOfWeek == DayOfWeek.Sunday)
                        {
                            if (Day3 != Day4.AddDays(1))
                            {
                                Day4 = Day4.AddDays(1);
                            }
                        }
                    }
                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        if (flag == 1)
                        {
                            Day5 = Recievedate.AddDays(6);
                            flag = 1;
                        }
                        else
                        {
                            Day5 = Recievedate.AddDays(3);
                            flag = 0;
                        }

                    }
                    else
                    {
                        if (Day2 == Recievedate.AddDays(4) || Day3 == Recievedate.AddDays(4) || Day4 == Recievedate.AddDays(4))
                        {
                            Day5 = Recievedate.AddDays(6);
                        }
                        else
                        {
                            Day5 = Recievedate.AddDays(4);
                        }
                        if (Day5.DayOfWeek == DayOfWeek.Saturday)
                        {
                            Day5 = Day5.AddDays(2);
                        }
                        if (Day5.DayOfWeek == DayOfWeek.Sunday)
                        {
                            if (Day4 != Day5.AddDays(1))
                            {
                                Day5 = Day5.AddDays(1);
                            }
                        }


                    }
                    if (Recievedate.DayOfWeek == DayOfWeek.Friday)
                    {
                        if (flag == 1)
                        {
                            Day6 = Recievedate.AddDays(7);
                            flag = 1;
                        }
                        else
                        {
                            Day6 = Recievedate.AddDays(3);
                            flag = 0;
                        }

                    }
                    else
                    {
                        Day6 = Recievedate.AddDays(5);
                    }

                    extraDay = Day6;
                    //loop 5 times for day1,day2,day3,day4,day5
                    for (int k = 0; k < GroupData.Rows.Count; k++)
                    {
                        Processdate = Convert.ToDateTime(GroupData.Rows[k]["ProcessDate"]);

                        if (Processdate == Recievedate)
                        {
                            Day1Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day1Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);
                        }

                        //if (GroupData.Rows.Count > 1)
                        if (Processdate == Day2)
                        {
                            Day2Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day2Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);
                        }


                        if (Processdate == Day3)
                        {
                            Day3Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day3Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);

                        }


                        if (Processdate == Day4)
                        {
                            Day4Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day4Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);

                        }


                        if (Processdate == Day5)
                        {
                            Day5Invoice = Convert.ToInt32(GroupData.Rows[k]["InvoiceCount"]);
                            Day5Percenatage = Convert.ToDecimal(GroupData.Rows[k]["Day1"]);

                        }


                    }



                    SqlQuery1 = "INSERT INTO REPORT VALUES('" + Recievedate + "'," + TotalInvoices + "," + Pending + "," + Issue + "," + Day1Invoice + "," + Day1Percenatage + "," + Day2Invoice + "," + Day2Percenatage + "," + Day3Invoice + "," + Day3Percenatage + "," + Day4Invoice + "," + Day4Percenatage + "," + Day5Invoice + "," + Day5Percenatage + "," + Day6Invoice + "," + Day6Percenatage + ")";
                    using (SqlConnection con = new SqlConnection(conString1))
                    {
                        con.Open();
                        SqlCommand cmd = new SqlCommand(SqlQuery1, con);
                        cmd.ExecuteNonQuery();
                    }

                }
            }


            //get 6plus day report

            // DataTable PlusDay_dt = GetDataEDI(";with CTE_Input1 as (select cast(FI_CreatedOn as date ) as RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate,IND_InvoiceNo  from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_Status not in ('IP_Asg','IP_Inp','IP_Issue')and IND_IP_Processed_By !='null' and FI_Source in ('EDI') ),CTE_Input2 as (select cast(FI_CreatedOn as date ) as RecieveDate, sum(cast(IND_InvoiceNo as int)) as TotalInvoices  from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and FI_Source in ('EDI') group by cast(FI_CreatedOn as date )),CTE_Input3 as(select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input2.TotalInvoices as TotalInvoices,sum(cast(CTE_Input1.IND_InvoiceNo as int)) as InvoiceCount  from CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate  group by CTE_Input1.RecieveDate,CTE_Input1.ProcessDate,TotalInvoices),CTE_Input4 as(select RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,case when InvoiceCount<>0 then (CAST(InvoiceCount as float)/CAST(TotalInvoices as float)*100.0) else 100 END as Accuracy from CTE_Input3 ), CTE_Input5 as (select RecieveDate,sum(InvoiceCount)as Pending from CTE_Input4 group by RecieveDate),CTE_Input6 as( select  CTE_Input4.RecieveDate as RecieveDate,ProcessDate,TotalInvoices,InvoiceCount, Cast(round(Accuracy,2) as decimal(10,2)) as Day1,CTE_Input5.Pending as Pending from CTE_Input5 join CTE_Input4 on CTE_Input4.RecieveDate=CTE_Input5.RecieveDate) select  RecieveDate,TotalInvoices,InvoiceCount as '6PlusDay',Day1 as '6PlusDay%' from CTE_Input6 WHERE ProcessDate >= DATEADD(day,5, RecieveDate) order by RecieveDate");

            DataTable PlusDay_dt = GetDataEDI(";with CTE_Input1 as (select cast(FI_CreatedOn as date ) as RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate,IND_InvoiceNo  from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_Status not in ('IP_Asg','IP_Inp','IP_Issue')and IND_IP_Processed_By !='null' and FI_Source in ('EDI') ),CTE_Input2 as (select cast(FI_CreatedOn as date ) as RecieveDate, sum(cast(IND_InvoiceNo as int)) as TotalInvoices  from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and FI_Source in ('EDI') group by cast(FI_CreatedOn as date )),CTE_Input3 as(select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input2.TotalInvoices as TotalInvoices,sum(cast(CTE_Input1.IND_InvoiceNo as int)) as InvoiceCount  from CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate  group by CTE_Input1.RecieveDate,CTE_Input1.ProcessDate,TotalInvoices),CTE_Input4 as(select RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,case when InvoiceCount<>0 then (CAST(InvoiceCount as float)/CAST(TotalInvoices as float)*100.0) else 100 END as Accuracy from CTE_Input3 ), CTE_Input5 as (select RecieveDate,sum(InvoiceCount)as Pending from CTE_Input4 group by RecieveDate),CTE_Input6 as( select  CTE_Input4.RecieveDate as RecieveDate,ProcessDate,TotalInvoices,InvoiceCount, Cast(round(Accuracy,2) as decimal(10,2)) as Day1,CTE_Input5.Pending as Pending from CTE_Input5 join CTE_Input4 on CTE_Input4.RecieveDate=CTE_Input5.RecieveDate) select  RecieveDate,TotalInvoices,InvoiceCount as '6PlusDay',Day1 as '6PlusDay%' from CTE_Input6 WHERE ProcessDate >= (select(DATEADD(day, (5 % 5) + CASE ((@@DATEFIRST + DATEPART(weekday, RecieveDate) + (5 % 5)) % 7) WHEN 0 THEN 2 WHEN 1 THEN 1 ELSE 0 END, DATEADD(week, (5 / 5), RecieveDate))) as RecieveDate) order by RecieveDate");
            for (int z = 0; z < PlusDay_dt.Rows.Count; z++)
            {
                int plusdayCount = Convert.ToInt32(PlusDay_dt.Rows[z]["6PlusDay"]);
                decimal plusdayPercentage = Convert.ToInt32(PlusDay_dt.Rows[z]["6PlusDay%"]);
                DateTime Recievedate = Convert.ToDateTime(PlusDay_dt.Rows[z]["RecieveDate"]);

                //update 6plus day data.
                SqlQuery1 = " update Report set Day6Count=" + plusdayCount + ", Day6Percentage='" + plusdayPercentage + "' where RECEIVEDATE='" + Recievedate + "'";
                using (SqlConnection con = new SqlConnection(conString1))
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(SqlQuery1, con);
                    cmd.ExecuteNonQuery();
                }
            }

        }

        private string getDelayHTMLEDI(DataTable dt)
        {

            myBuilderEDI = new StringBuilder();

            myBuilderEDI.Append("<table border='1' cellpadding='5' cellspacing='0' ");
            myBuilderEDI.Append("style='border: solid 1px Silver; font-size: x-small;'>");

            myBuilderEDI.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in dt.Columns)
            {
                myBuilderEDI.Append("<td align='left' valign='top' bgcolor='#3F88F6'>");
                myBuilderEDI.Append("<B />" + myColumn.ColumnName);
                myBuilderEDI.Append("</td>");
            }
            myBuilderEDI.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                if (myRow.ItemArray.Contains("Total"))
                {
                    myBuilderEDI.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilderEDI.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                        myBuilderEDI.Append("<B />" + myRow[myColumn.ColumnName].ToString());
                        myBuilderEDI.Append("</td>");
                    }
                    myBuilderEDI.Append("</tr>");
                }
                else
                {
                    myBuilderEDI.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilderEDI.Append("<td align='left' valign='top'>");
                        myBuilderEDI.Append(myRow[myColumn.ColumnName].ToString());
                        myBuilderEDI.Append("</td>");
                    }
                    myBuilderEDI.Append("</tr>");
                }

            }
            myBuilderEDI.Append("</table>");

            return myBuilderEDI.ToString();
        }

        private string getDelayHTML1EDI(DataTable dt)
        {

            myBuilder1EDI = new StringBuilder();

            myBuilder1EDI.Append("<table border='1' cellpadding='5' cellspacing='0' ");
            myBuilder1EDI.Append("style='border: solid 1px Silver; font-size: x-small;'>");

            myBuilder1EDI.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in dt.Columns)
            {
                myBuilder1EDI.Append("<td align='left' valign='top' bgcolor='#3F88F6'>");
                myBuilder1EDI.Append("<B />" + myColumn.ColumnName);
                myBuilder1EDI.Append("</td>");
            }
            myBuilder1EDI.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                if (myRow.ItemArray.Contains("Total"))
                {
                    myBuilder1EDI.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder1EDI.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                        myBuilder1EDI.Append("<B />" + myRow[myColumn.ColumnName].ToString());
                        myBuilder1EDI.Append("</td>");
                    }
                    myBuilder1EDI.Append("</tr>");
                }
                else
                {
                    myBuilder1EDI.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder1EDI.Append("<td align='left' valign='top'>");
                        myBuilder1EDI.Append(myRow[myColumn.ColumnName].ToString());
                        myBuilder1EDI.Append("</td>");
                    }
                    myBuilder1EDI.Append("</tr>");
                }

            }
            myBuilder1EDI.Append("</table>");

            return myBuilder1EDI.ToString();
        }

        private string getDelayHTML2EDI(DataTable dt)
        {

            myBuilder2EDI = new StringBuilder();

            myBuilder2EDI.Append("<table border='1' cellpadding='5' cellspacing='0' ");
            myBuilder2EDI.Append("style='border: solid 1px Silver; font-size: x-small;'>");

            myBuilder2EDI.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in dt.Columns)
            {
                myBuilder2EDI.Append("<td align='left' valign='top' bgcolor='#3F88F6'>");
                myBuilder2EDI.Append("<B />" + myColumn.ColumnName);
                myBuilder2EDI.Append("</td>");
            }
            myBuilder2EDI.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                if (myRow.ItemArray.Contains("Total"))
                {
                    myBuilder2EDI.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder2EDI.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                        myBuilder2EDI.Append("<B />" + myRow[myColumn.ColumnName].ToString());
                        myBuilder2EDI.Append("</td>");
                    }
                    myBuilder2EDI.Append("</tr>");
                }
                else
                {
                    myBuilder2EDI.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder2EDI.Append("<td align='left' valign='top'>");
                        myBuilder2EDI.Append(myRow[myColumn.ColumnName].ToString());
                        myBuilder2EDI.Append("</td>");
                    }
                    myBuilder2EDI.Append("</tr>");
                }

            }
            myBuilder2EDI.Append("</table>");

            return myBuilder2EDI.ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
           
            lblStart.Text = "Downloaded started";
           downloadObj.MoveDirectory();
           lblMove.Text = "Files Moved";
           downloadObj.copyDirectory();
           lblCopy.Text = "Files Copied";
           downloadObj.DataProcessing();
           MessageBox.Show("Downloaded successfully");
        }
    }
}
