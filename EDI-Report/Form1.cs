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
//using ClosedXML.Excel;
using System.Text;
using System.IO;
using System.Threading;
using ClosedXML.Excel;
namespace EDI_Report
{
    public partial class Form1 : Form
    {
        StringBuilder myBuilder = new StringBuilder();
        StringBuilder myBuilder1 = new StringBuilder();
        StringBuilder myBuilder2 = new StringBuilder();
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
            DayOfWeek dayToday = DateTime.Now.DayOfWeek;
            List<DayOfWeek> partyDays = new List<DayOfWeek> {
                     DayOfWeek.Saturday, DayOfWeek.Sunday
                              };

            if (partyDays.Contains(dayToday))
            {
                //do nothing.
            }
            else
            {
                TimeSpan start1 = new TimeSpan(20, 30, 0);// New TimeSpan(6, 45, 0)
                TimeSpan end1 = new TimeSpan(20, 31, 0);
                TimeSpan now = TimeSpan.Parse(DateTime.Now.TimeOfDay.ToString("hh\\:mm\\:ss"));

                int i = 0;
                if (now == start1)
                {
                    if (i == 0)
                    {
                        WeeklyReport();
                        i++;
                    }
                }
                if (now == end1)
                {
                    i = 0;
                }
            }

            if ((DayOfWeek.Saturday == DayOfWeek.Saturday))
            {
                //do nothing
            }
            else
            {
            }

            //    BackgroundWorker helperBW = sender as BackgroundWorker;
            //    int arg = (int)e.Argument;
            //    e.Result = BackgroundProcessLogicMethod(helperBW, arg);
            //    if (helperBW.CancellationPending)
            //    {
            //        e.Cancel = true;
            //    }
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


        string conString;
        string SqlQuery;
        DateTime From;
        DateTime To;
        private void WeeklyReport()
        {
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



            //schedular
            //DateTime From = Convert.ToDateTime(DateTime.Today.AddDays(-7));
            //DateTime To = Convert.ToDateTime(DateTime.Now);

            //DataTable pending = GetData("select  ReceivedDate, Pending from (select  cast(FI_CreatedOn as date )as ReceivedDate,sum([IP_Asg]) as Pending from (select * from (select FI_CreatedOn ,IND_Status  from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID) src  pivot (count(IND_Status)  for IND_Status in ([IP_Asg]) ) piv) aaa where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' group by cast(FI_CreatedOn as date))a order by ReceivedDate");
            DataTable pending = GetData("select ReceivedDate, Issue+Invoicesassigned+InvoiceInProgress+invoiceUnassigned as pending from(select ReceivedDate,TotalInvoicesReceived,Invoicesassigned,InvoiceInProgress,(TotalInvoicesReceived - (Invoicesassigned+InvoiceInProgress+QCUnassinged+QCAssigned+QCInprogress+QCCompleted+Duplicate+Issue+EDI+DNP+Expedite+Statement)) as invoiceUnassigned,(QCUnassinged+QCAssigned+QCInprogress+QCCompleted) as InvoiceCompleted,Duplicate,Issue,EDI,DNP,Expedite from (select cast(FI_CreatedOn as date )as ReceivedDate,count (FI_OriginalName) as TotalInvoicesReceived,sum([IP_Asg]) as Invoicesassigned,sum([IP_Inp]) as InvoiceInProgress,sum ([QC_Idle]) as QCUnassinged,sum([QC_Asg])as QCAssigned,sum([QC_Inp])as QCInprogress,sum([QC_Comp])as QCCompleted,sum([Duplicate])as Duplicate,sum([IP_Issue])as Issue, SUM([EDI]) as EDI, SUM([DNP]) as DNP, SUM([Expedite]) as Expedite,SUM([Statement]) as Statement from (select * from (select FI_OriginalName,FI_CreatedOn ,IND_Status,IND_FI  from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID ) src   pivot (count(IND_Status)  for IND_Status in ([IP_Asg],[QC_Inp],[QC_Idle],[IP_Inp],[QC_Comp],[QC_Asg],[Duplicate],[IP_Issue],[EDI],[DNP],[Expedite],[Statement]) ) piv) aaa   where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and    cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' group by cast(FI_CreatedOn as date ))a)b ");
            DataTable dt = new DataTable();

            //dt = GetData(";with CTE_Input1 as ( select RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate, count(FI_OriginalName) as InvoiceCount from  (select cast(FI_CreatedOn as date ) as RecieveDate, IND_IP_ModifiedOn,FI_OriginalName from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID   where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and IND_Status not in ('IP_Asg','IP_Issue')) A group by cast(IND_IP_ModifiedOn as date),RecieveDate),CTE_Input2 as (select cast(FI_CreatedOn as date )as RecieveDate, count(FI_OriginalName) as TotalInvoices from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID  where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' group by cast(FI_CreatedOn as date )),CTE_Input3 as(SELECT RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,Cast(round(Accuracy,2) as decimal(10,2)) as Day1 FROM (select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input1.InvoiceCount as InvoiceCount,CTE_Input2.TotalInvoices as TotalInvoices, case when CTE_Input1.InvoiceCount<>0 then (CAST(CTE_Input1.InvoiceCount as float)/CAST(CTE_Input2.TotalInvoices as float)*100.0) else 100 END as Accuracy from  CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate where ProcessDate<= DATEADD(day,DATEDIFF(day,0,CTE_Input1.RecieveDate),7) )B ) " +
            //            "select * from CTE_Input3");

            dt = GetData(";with CTE_Input1 as ( select RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate, count(FI_OriginalName) as InvoiceCount from  (select cast(FI_CreatedOn as date ) as RecieveDate, IND_IP_ModifiedOn,FI_OriginalName from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID   where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and IND_Status not in ('IP_Asg','IP_Inp','IP_Issue')) A group by cast(IND_IP_ModifiedOn as date),RecieveDate),CTE_Input2 as (select cast(FI_CreatedOn as date )as RecieveDate, count(FI_OriginalName) as TotalInvoices from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID  where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' group by cast(FI_CreatedOn as date )),CTE_Input3 as(SELECT RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,Cast(round(Accuracy,2) as decimal(10,2)) as Day1 FROM (select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input1.InvoiceCount as InvoiceCount,CTE_Input2.TotalInvoices as TotalInvoices, case when CTE_Input1.InvoiceCount<>0 then (CAST(CTE_Input1.InvoiceCount as float)/CAST(CTE_Input2.TotalInvoices as float)*100.0) else 100 END as Accuracy from  CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate )B ) " +
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
                        }
                    }

                    DateTime Day2 = Recievedate.AddDays(1);
                    DateTime Day3 = Recievedate.AddDays(2);
                    DateTime Day4 = Recievedate.AddDays(3);
                    DateTime Day5 = Recievedate.AddDays(4);
                    DateTime Day6 = Recievedate.AddDays(5);

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



                    SqlQuery = "INSERT INTO REPORT VALUES('" + Recievedate + "'," + TotalInvoices + "," + Pending + "," + Day1Invoice + "," + Day1Percenatage + "," + Day2Invoice + "," + Day2Percenatage + "," + Day3Invoice + "," + Day3Percenatage + "," + Day4Invoice + "," + Day4Percenatage + "," + Day5Invoice + "," + Day5Percenatage + "," + Day6Invoice + "," + Day6Percenatage + ")";
                    using (SqlConnection con = new SqlConnection(conString))
                    {
                        con.Open();
                        SqlCommand cmd = new SqlCommand(SqlQuery, con);
                        cmd.ExecuteNonQuery();
                    }

                }
            }


            //get 6plus day report
            DataTable PlusDay_dt = GetData(";with CTE_Input1 as ( select RecieveDate,cast(IND_IP_ModifiedOn as date) as ProcessDate, count(FI_OriginalName) as InvoiceCount from (select cast(FI_CreatedOn as date ) as RecieveDate, IND_IP_ModifiedOn,FI_OriginalName from dbo.EMSDB_FileInfo  left join dbo.EMSDB_InvDetails on IND_FI=FI_ID   where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' and IND_IP_Processed_By !='null' and IND_Status not in ('IP_Asg','IP_Issue')) A group by cast(IND_IP_ModifiedOn as date),RecieveDate),CTE_Input2 as (select cast(FI_CreatedOn as date )as RecieveDate,count(FI_OriginalName) as TotalInvoices from dbo.EMSDB_FileInfo   left join dbo.EMSDB_InvDetails on IND_FI=FI_ID  where cast(FI_CreatedOn as date ) >='" + From.ToString("yyyy-MM-dd") + "'  and cast(FI_CreatedOn as date ) <='" + To.ToString("yyyy-MM-dd") + "' group by cast(FI_CreatedOn as date )),CTE_Input3 as (SELECT RecieveDate,ProcessDate,TotalInvoices,InvoiceCount,Cast(round(Accuracy,2) as decimal(10,2)) as Day1 FROM (select CTE_Input1.RecieveDate as RecieveDate,CTE_Input1.ProcessDate as ProcessDate,CTE_Input1.InvoiceCount as InvoiceCount,CTE_Input2.TotalInvoices as TotalInvoices, case when CTE_Input1.InvoiceCount<>0 then (CAST(CTE_Input1.InvoiceCount as float)/CAST(CTE_Input2.TotalInvoices as float)*100.0) else 100 END as Accuracy from CTE_Input1 join CTE_Input2 on CTE_Input1.RecieveDate=CTE_Input2.RecieveDate )B ) select  RecieveDate,TotalInvoices,PlusDay as '6PlusDay',Cast(round(Accuracy,2) as decimal(10,2)) as '6PlusDay%' from (select RecieveDate,TotalInvoices,PlusDay,case when PlusDay<>0 then (CAST(PlusDay as float)/CAST(TotalInvoices as float)*100.0) else 100 END as Accuracy  from (select RecieveDate,TotalInvoices,sum(InvoiceCount) as PlusDay from CTE_Input3 WHERE ProcessDate >= DATEADD(day,5, RecieveDate)group by RecieveDate,TotalInvoices " +
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
            DataTable pendingReport_dt = GetData("select convert(varchar, RECEIVEDATE, 120) as DateAdded,Pending from REPORT where Pending<>0");
            //generate SLA report.
            DataTable SLA_dt = GetData("select convert(varchar, RECEIVEDATE, 120) as DateAdded,TotalInvoice,ProcessedInvoices as Met_GS_SLA,Convert(varchar(10),(CAST(round(Percentage,2) as decimal(10,2))))+'%' as GS_SLA from (select RECEIVEDATE,TotalInvoice,sum(Day1Count+Day2Count+Day3Count) as ProcessedInvoices,case when sum(Day1Count+Day2Count+Day3Count)<>0 then (CAST(sum(Day1Count+Day2Count+Day3Count) as float)/CAST(TotalInvoice as float)*100.0) else 100 END as Percentage from dbo.REPORT  group by RECEIVEDATE,TotalInvoice ) as a order by RECEIVEDATE");



            //Set DataTable Name which will be the name of Excel Sheet.
            weeklyReport_dt.TableName = "EMSDB Weekly Report";
            pendingReport_dt.TableName = "EMSDB Pending Report";
            SLA_dt.TableName = "EMSDB SLA Report";

            //Create a New Workbook.
            using (XLWorkbook wb = new XLWorkbook())
            {
                //Add the DataTable as Excel Worksheet.
                wb.Worksheets.Add(weeklyReport_dt);
                wb.Worksheets.Add(pendingReport_dt);
                wb.Worksheets.Add(SLA_dt);
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

                    message1.To.Add(("Manjunath.Jayaram@tangoe.com"));
                    message1.To.Add(("Sriram.Krishnan@tangoe.com"));
                    message1.To.Add(("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(("Johwessly.Chennaiah@tangoe.com"));

                    message1.To.Add(("Rashmi.Ahuja@tangoe.com"));
                    //message1.To.Add(("rich.lena@tangoe.com"));
                    message1.To.Add(("melissa.guarracino@tangoe.com"));

                    message1.CC.Add(("Sleema.Joseph@tangoe.com"));
                    message1.CC.Add(("namohar.m@tangoe.com"));
                    //  message1.To.Add(("namohar.m@tangoe.com"));

                    System.Net.Mail.Attachment attachment;

                    message1.Attachments.Add(new Attachment(new MemoryStream(bytes), "EDI-EMSDBWeeklyReport.xlsx"));


                    message1.Subject = "Global Support EMS Invoice Processing EOB Report for  " + To.ToString("yyyy-MM-dd");

                    //   message1.Body = string.Format("<h3>Hi Team,<h3>" +
                    //"<br /> EMSDB Weekly Report for " + From.ToString("M/d/yyyy") + " to " + To.ToString("M/d/yyyy") + ". Please review and let us know if you have any clarifications.<br /><br />Note : This is System Generated mail. please send mail to namohar.m@tangoe.com if you have any issues.<br /><br /> Thank You.<br />EMSDB Team<br /> ");
                    // message1.Body = getHTML(weeklyReport_dt);

                    getDelayHTML1(pendingReport_dt);
                    getDelayHTML(weeklyReport_dt);
                    getDelayHTML2(SLA_dt);

                    StringBuilder sb = new StringBuilder();

                    sb.AppendLine("<h3>Hi Team,<h3>");

                    sb.AppendLine("");
                    sb.AppendLine(string.Format("<br /> EMSDB-EDI Weekly Report for " + From.ToString("yyyy-MM-dd") + " to " + To.ToString("yyyy-MM-dd")));
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
                    sb.AppendLine("");//here I want the data to display in table format
                    sb.AppendLine(string.Format("Please review and let us know if you have any clarifications.<br /><br /><P> Note : This is system generated mail. please send mail to namohar.m@tangoe.com and tangoe-devops@tangoe.com if you have any issues.</p><br /><br /> Thank You.<br />EMSDB Team<br /> "));
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
            // string conString = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;
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
    }
}
