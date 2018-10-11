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
namespace VodafoneEDI_Download
{
    public partial class Form1 : Form
    {
        //Testing
        //SqlConnection con = new SqlConnection("Data Source=10.80.20.61,1433;Initial Catalog=EMSDB_QA;User ID=opsdev;Password=opsdev@123;");
        //Production
        SqlConnection con = new SqlConnection("Data Source=BLRPRODRTM\\RTM_PROD_BLR;Initial Catalog=WorkFlowManagerDB;User ID=sa;Password=Prodrtm@123;");


        string Path;
        DateTime CurrentDate = DateTime.Now;
        string dtnow;
        DataTable dt = new DataTable();
        public Form1()
        {
            InitializeComponent();
            //DataProcessing();
        }
        private DataTable GetClients()
        {
            DataTable dt = new DataTable();
            using (SqlDataAdapter ada = new SqlDataAdapter("select * from Vodafone_Clients where C_Status=1", con))
                ada.Fill(dt);
            return dt;
        }
        public void DataProcessing()
        {

            DataTable dt = new DataTable();
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("Path", typeof(string));


            DataRow dtrow = dt.NewRow();    // Create New Row            //Bind Data to Columns
            dtrow["Type"] = "EDI";
          
            dtrow["Path"] = @"\\10.80.20.251\InvoicesRepository\Vodafone\Production\EDI";
            dt.Rows.Add(dtrow);
            //dtrow = dt.NewRow();
            //dtrow["Type"] = "PDF Portal";
            //dtrow["Path"] = @"\\10.80.20.251\InvoicesRepository\Vodafone\Production\PDF Portal Downloads";
            //dt.Rows.Add(dtrow);
            //dtrow = dt.NewRow();
            //dtrow["Type"] = "Scanned Mail";
            //dtrow["Path"] = @"\\10.80.20.251\InvoicesRepository\Vodafone\Production\Scanned Mail";
            //dt.Rows.Add(dtrow);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string path = dt.Rows[i]["Path"].ToString();
                string Type = dt.Rows[i]["Type"].ToString();
                DataInsertion(path, Type);
            }

            MessageBox.Show("Data successfully inserted");
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
                                    cmd.Parameters.AddWithValue("@FI_FileName", FullPath);
                                    cmd.Parameters.AddWithValue("@FI_ContentType", "text/plain");
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
            //Testing
            //using (SqlConnection con = new SqlConnection("Data Source=10.80.20.61,1433;Initial Catalog=EMSDB_QA;User ID=opsdev;Password=opsdev@123;"))
            
            //Production
            using (SqlConnection con = new SqlConnection("Data Source=BLRPRODRTM\\RTM_PROD_BLR;Initial Catalog=WorkFlowManagerDB;User ID=sa;Password=Prodrtm@123;"))
            {
                da = new SqlDataAdapter(ssql, con);
                da.Fill(table);
            }
            return table;
        }

        private void btnDownload_Click(object sender, EventArgs e)
        {
            DataProcessing();
        }
    }
}
