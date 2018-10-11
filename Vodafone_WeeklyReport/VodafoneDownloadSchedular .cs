using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace Vodafone_WeeklyReport
{
    class VodafoneDownloadSchedular
    {
        //Testing
        //SqlConnection con = new SqlConnection("Data Source=10.80.20.61,1433;Initial Catalog=EMSDB_QA;User ID=opsdev;Password=opsdev@123;");
        //Production
        SqlConnection con = new SqlConnection("Data Source=BLRPRODRTM\\RTM_PROD_BLR;Initial Catalog=WorkFlowManagerDB;User ID=sa;Password=Prodrtm@123;");
        string Path;
        DateTime CurrentDate = DateTime.Now;
        string dtnow;
        DataTable dt = new DataTable();
        private DataTable GetClients()
        {
            DataTable dt = new DataTable();
            using (SqlDataAdapter ada = new SqlDataAdapter("select * from Vodafone_Clients where C_Status=1 and C_Code not in ('3M','CSC','Exxon','GLENCORE','MMC','HELLO WORLD','Honda','HP','TDAmeritrade')", con))
                ada.Fill(dt);
            return dt;
        }

        public void MoveDirectory()
        {
            try
            {
                dtnow = CurrentDate.ToString("MM dd yyyy");
                dt = GetClients();              

                //Digitizing OutLook
                foreach (DataRow dr in dt.Rows)
                {
                    string client = dr["C_Name"].ToString();
                    //Testing

                    //string sourceDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\Digitizing OutLook\" + client + "\\Monthly Invoices";
                    //string destDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\Digitizing OutLook\" + client + "\\Monthly Invoices\\Archive\\" + dtnow;

                    //Production
                    string sourceDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Digitizing OutLook\" + client + "\\Monthly Invoices";
                    string destDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Digitizing OutLook\" + client + "\\Monthly Invoices\\Archive\\" + dtnow;
                    Move(sourceDirName, destDirName);
                }

                //PDF Portal
                foreach (DataRow dr in dt.Rows)
                {
                    string client = dr["C_Name"].ToString();
                    //TESTING
                    //string sourceDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\PDF Portal Downloads\" + client;
                    //string destDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\PDF Portal Downloads\" + client + "\\Archive\\" + dtnow;

                    //Production
                    string sourceDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\PDF Portal Downloads\" + client;
                    string destDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\PDF Portal Downloads\" + client + "\\Archive\\" + dtnow;
                    Move(sourceDirName, destDirName);
                }

                //Scanned file
                foreach (DataRow dr in dt.Rows)
                {
                    string client = dr["C_Name"].ToString();

                    if (client == "Tyco")
                    {
                        string client1 = client;
                    }
                    //Testing
                    //string sourceDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\Scanned Mail\" + client + "\\Scanned Invoice";
                    //string destDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\Scanned Mail\" + client + "\\Scanned Invoice\\Archive\\" + dtnow;

                    //Production
                    string sourceDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Scanned Mail\" + client + "\\Scanned Invoice";
                    string destDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Scanned Mail\" + client + "\\Scanned Invoice\\Archive\\" + dtnow;
                    Move(sourceDirName, destDirName);
                }
                //DNR file
                foreach (DataRow dr in dt.Rows)
                {
                    string client = dr["C_Name"].ToString();

                    if (client == "Tyco")
                    {
                        string client1 = client;
                    }
                    //Testing
                    //string sourceDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\Scanned Mail\" + client + "\\Scanned Invoice";
                    //string destDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\Scanned Mail\" + client + "\\Scanned Invoice\\Archive\\" + dtnow;

                    //Production
                    string sourceDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\DNR Files\" + client;
                    string destDirName = @"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\DNR Files\" + client + "\\Archive\\" + dtnow;
                    Move(sourceDirName, destDirName);
                }

            }
            catch (Exception ex)
            {
                throw ex;

            }

        }

        public void Move(string sourceDirName, string destDirName)
        {

            //dtnow = CurrentDate.ToString("MM dd yyyy");
            // dt = GetClients();

            // foreach (DataRow dr in dt.Rows)
            // {
            //     string client = dr["C_Name"].ToString();
            //string sourceDirName = @"\\10.80.20.251\EDIRepository\Vodafone DB - Test\Data In Life\Digitizing OutLook\"+client+"\\Monthly Invoices";
            //string destDirName = @"\\10.80.20.251\EDIRepository\Vodafone DB - Test\Data In Life\Digitizing OutLook\"+client+"\\Monthly Invoices\\Archive\\" + dtnow;
            DirectoryInfo destinationDir = new DirectoryInfo(destDirName);
            if (!destinationDir.Exists)
            {
                destinationDir.Create();
            }
            //try
            //{
            //    Directory.Move(sourceDirName, destDirName);
            //}
            //catch (IOException exp)
            //{
            //   Console.WriteLine(exp.Message);
            //}
            //  string foldername = @"\\10.80.20.251\OPSDevRepository\NAMOHAR\Archive";
            string foldername = sourceDirName;
            DirectoryInfo di = new DirectoryInfo(foldername);
            if (di.Exists)
            {
                //DirectoryInfo[] dirInfo = di.GetDirectories();
                //foreach (DirectoryInfo file in dirInfo)
                //{
                //    string folder1 = file.FullName.ToString();
                //    Path = folder1;

                //DirectoryInfo dii = new DirectoryInfo(Path);
                DirectoryInfo dii = new DirectoryInfo(foldername);
                if (dii.Exists)
                {
                    FileInfo[] subdirInfo = dii.GetFiles();

                    foreach (FileInfo files in subdirInfo)
                    {
                        string Filename = files.Name;
                        var result = Filename.Substring(Filename.Length - 3);
                        if (result != ".db")
                        {
                            File.Move(sourceDirName + "\\" + Filename, destDirName + "\\" + Filename);
                        }
                        else
                        {
                        }
                    }
                }
                //}
            }

        }

        public void copyDirectory()
        {
            //get the client list
            dt = GetClients();
            //Digitizing
            foreach (DataRow dr in dt.Rows)
            {
                //get the client name
                string client = dr["C_Name"].ToString();
                //Testing
                //DirectoryInfo sourceDir = new DirectoryInfo(@"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\Digitizing OutLook\" + client + "\\Monthly Invoices\\Archive\\" + dtnow);
                //DirectoryInfo destinationDir = new DirectoryInfo(@"\\10.80.20.251\InvoicesRepository\Vodafone\Production\Digitizing OutLook\" + client + "\\" + dtnow);

                //Production
                DirectoryInfo sourceDir = new DirectoryInfo(@"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Digitizing OutLook\" + client + "\\Monthly Invoices\\Archive\\" + dtnow);
                DirectoryInfo destinationDir = new DirectoryInfo(@"\\10.80.20.251\InvoicesRepository\Vodafone\Production\Digitizing OutLook\" + client + "\\" + dtnow);
                Copy(sourceDir, destinationDir);
            }

            //PDF Portal
            foreach (DataRow dr in dt.Rows)
            {
                //get the client name
                string client = dr["C_Name"].ToString();
                //Testing
                //DirectoryInfo sourceDir = new DirectoryInfo(@"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\PDF Portal Downloads\" + client + "\\Archive\\" + dtnow);
                //DirectoryInfo destinationDir = new DirectoryInfo(@"\\10.80.20.251\InvoicesRepository\Vodafone\Production\PDF Portal Downloads\" + client + "\\" + dtnow);

                //Production
                DirectoryInfo sourceDir = new DirectoryInfo(@"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\PDF Portal Downloads\" + client + "\\Archive\\" + dtnow);
                DirectoryInfo destinationDir = new DirectoryInfo(@"\\10.80.20.251\InvoicesRepository\Vodafone\Production\PDF Portal Downloads\" + client + "\\" + dtnow);
                Copy(sourceDir, destinationDir);
            }

            //Scanned Mail
            foreach (DataRow dr in dt.Rows)
            {
                //get the client name.
                string client = dr["C_Name"].ToString();
                //Testing
                //DirectoryInfo sourceDir = new DirectoryInfo(@"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\Scanned Mail\" + client + "\\Scanned Invoice\\Archive\\" + dtnow);
                //DirectoryInfo destinationDir = new DirectoryInfo(@"\\10.80.20.251\InvoicesRepository\Vodafone\Production\Scanned Mail\" + client + "\\" + dtnow);

                //Production
                DirectoryInfo sourceDir = new DirectoryInfo(@"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Scanned Mail\" + client + "\\Scanned Invoice\\Archive\\" + dtnow);
                DirectoryInfo destinationDir = new DirectoryInfo(@"\\10.80.20.251\InvoicesRepository\Vodafone\Production\Scanned Mail\" + client + "\\" + dtnow);
                Copy(sourceDir, destinationDir);
            }

            //DNR Files
            foreach (DataRow dr in dt.Rows)
            {
                //get the client name.
                string client = dr["C_Name"].ToString();
                //Testing
                //DirectoryInfo sourceDir = new DirectoryInfo(@"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\Data In Life - TEST\Scanned Mail\" + client + "\\Scanned Invoice\\Archive\\" + dtnow);
                //DirectoryInfo destinationDir = new DirectoryInfo(@"\\10.80.20.251\InvoicesRepository\Vodafone\Production\Scanned Mail\" + client + "\\" + dtnow);

                //Production
                DirectoryInfo sourceDir = new DirectoryInfo(@"\\dc2bkp03.corp.tangoe.com\shared$\Data In Life\DNR Files\" + client + "\\Archive\\" + dtnow);
                DirectoryInfo destinationDir = new DirectoryInfo(@"\\10.80.20.251\InvoicesRepository\Vodafone\Production\DNR\" + client + "\\" + dtnow);
                Copy(sourceDir, destinationDir);
            }
        }

        public void Copy(DirectoryInfo source, DirectoryInfo destination)
        {
            //if file is not created , here we can create programetically. 
            if (!destination.Exists)
            {
                destination.Create();
            }

            if (!source.Exists)
            {
                return;
            }

            // Copy all files.
            FileInfo[] files = source.GetFiles();
            //if files not exist do nothing
            if (files.Length == 0)
            {
                // do nothing
            }
            foreach (FileInfo file in files)
            {
                string Fname = destination.FullName + "//" + file.Name;
                //if file already exist skip the file and move to next loop.
                if (File.Exists(Fname))
                {
                    //do nothing skip it.
                }
                else
                {
                    file.CopyTo(System.IO.Path.Combine(destination.FullName, file.Name));
                }
            }

            // Process subdirectories.
            DirectoryInfo[] dirs = source.GetDirectories();
            foreach (DirectoryInfo dir in dirs)
            {
                // Get destination directory.
                string destinationDir = System.IO.Path.Combine(destination.FullName, dir.Name);

                // Call CopyDirectory() recursively.
                Copy(dir, new DirectoryInfo(destinationDir));
            }

        }
        
        public void DataProcessing()
        {

            DataTable dt = new DataTable();
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("Path", typeof(string));


            DataRow dtrow = dt.NewRow();    // Create New Row //Bind Data to Columns
            dtrow["Type"] = "Digitizing OutLook";
            dtrow["Path"] = @"\\10.80.20.251\InvoicesRepository\Vodafone\Production\Digitizing OutLook";
            dt.Rows.Add(dtrow);
            dtrow = dt.NewRow();
            dtrow["Type"] = "PDF Portal";
            dtrow["Path"] = @"\\10.80.20.251\InvoicesRepository\Vodafone\Production\PDF Portal Downloads";
            dt.Rows.Add(dtrow);
            dtrow = dt.NewRow();
            dtrow["Type"] = "Scanned Mail";
            dtrow["Path"] = @"\\10.80.20.251\InvoicesRepository\Vodafone\Production\Scanned Mail";
            dt.Rows.Add(dtrow);

            dtrow = dt.NewRow();
            dtrow["Type"] = "DNR";
            dtrow["Path"] = @"\\10.80.20.251\InvoicesRepository\Vodafone\Production\DNR";
            dt.Rows.Add(dtrow);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string path = dt.Rows[i]["Path"].ToString();
                string Type = dt.Rows[i]["Type"].ToString();
                DataInsertion(path, Type);
            }
        }
        //Data insertion 
        private void DataInsertion(string Location, string Type)
        {
            dtnow = CurrentDate.ToString("MM dd yyyy");
            string result = GetID();
            int id = Convert.ToInt32(result) + 1;
            dt = GetClients();

            foreach (DataRow dr in dt.Rows)
            {
                string foldername = string.Empty;
                string _originalFoldername = string.Empty;
                string filesource = string.Empty;
                string filedest = string.Empty;
                string _originalFiledest = string.Empty;
                string destination = string.Empty;
                string _originalDest = string.Empty;
                string Year = DateTime.Now.Year.ToString();
                string month = DateTime.Now.ToString("MMMM");


                foldername = Location + "\\" + dr["C_Name"].ToString();

                DirectoryInfo di = new DirectoryInfo(foldername);
                if (di.Exists)
                {
                    ////loop through folders *********
                    //DirectoryInfo[] dirInfo = di.GetDirectories();
                    //foreach (DirectoryInfo file in dirInfo)
                    //{
                    // string folder1 = file.FullName.ToString();                    
                    Path = foldername + "\\" + dtnow;
                    DirectoryInfo dii = new DirectoryInfo(Path);
                    if (dii.Exists)
                    {
                        FileInfo[] subdirInfo = dii.GetFiles();

                        foreach (FileInfo files in subdirInfo)
                        {
                            //you can manipulate the found files here
                            string Filename = files.Name;
                            //Checking File is exsit or not

                            DateTime filedate = CurrentDate;
                            string FullPath = Path + "\\" + Filename;
                            string Query = "select * from Vodafone_FileInfo where FI_Source='" + Type + "' and FI_OriginalName='" + Filename + "' and cast(FI_ReceiptDate as date )='" + filedate.ToString("yyyy/MM/dd") + "'";
                            DataTable DataExist = DBExecDataTable(Query);
                            if (DataExist.Rows.Count > 0)
                            {
                                //Do nothing
                            }
                            else
                            {
                                FileStream fs = new FileStream(Path + "\\" + Filename, FileMode.Open, FileAccess.ReadWrite);
                                BinaryReader br = new BinaryReader(fs);
                                Byte[] bytes = br.ReadBytes(Convert.ToInt32(fs.Length));
                                br.Close();
                                fs.Close();
                                using (SqlCommand cmd = new SqlCommand("Insert into Vodafone_FileInfo (FI_OriginalName, FI_ReceiptDate, FI_Source, FI_ClientCode, FI_FileName, FI_ContentType,FI_Data, FI_CreatedOn, FI_Status) Values (@FI_OriginalName, @FI_ReceiptDate, @FI_Source, @FI_ClientCode, @FI_FileName, @FI_ContentType, @FI_Data, @FI_CreatedOn, @FI_Status)", con))
                                {
                                    cmd.Parameters.AddWithValue("@FI_OriginalName", Filename);
                                    cmd.Parameters.AddWithValue("@FI_ReceiptDate", DateTime.Now);
                                    cmd.Parameters.AddWithValue("@FI_Source", Type);
                                    cmd.Parameters.AddWithValue("@FI_ClientCode", dr["C_Code"].ToString());
                                    //cmd.Parameters.AddWithValue("@FI_FileName", Filename);
                                    cmd.Parameters.AddWithValue("@FI_FileName", id + "_" + dr["C_Code"].ToString() + "_" + Filename);
                                    cmd.Parameters.AddWithValue("@FI_ContentType", "application/pdf");
                                    cmd.Parameters.AddWithValue("@FI_Data", bytes);
                                    cmd.Parameters.AddWithValue("@FI_CreatedOn", DateTime.Now);
                                    cmd.Parameters.AddWithValue("@FI_Status", "1");
                                    con.Open();
                                    cmd.ExecuteNonQuery();
                                    con.Close();
                                    id = id + 1;
                                }
                            }
                        }
                    }

                }
                //  }
            }

        }

        private string GetID()
        {
            SqlCommand cmd = new SqlCommand("Select ISNULL(MAX(FI_ID), 0) AS  FI_ID from Vodafone_FileInfo Order By FI_ID Desc", con);
            con.Open();
            string result = Convert.ToString(cmd.ExecuteScalar());
            con.Close();
            if (result == string.Empty)
            {
                return "";
            }
            else
            {
                return result;
            }
        }

        public DataTable DBExecDataTable(string ssql)
        {
            DataTable table = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter();
            //Tesing.
            //using (SqlConnection con = new SqlConnection("Data Source=10.80.20.61,1433;Initial Catalog=EMSDB_QA;User ID=opsdev;Password=opsdev@123;"))

            //Production.
            using (SqlConnection con = new SqlConnection("Data Source=BLRPRODRTM\\RTM_PROD_BLR;Initial Catalog=WorkFlowManagerDB;User ID=sa;Password=Prodrtm@123;"))
            {
                da = new SqlDataAdapter(ssql, con);
                da.Fill(table);
            }
            return table;
        }
    }
}
