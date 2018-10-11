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

namespace VodafoneUploader
{

    public partial class Form1 : Form
    {
        //Testing
        //SqlConnection con = new SqlConnection("Data Source=10.80.20.61,1433;Initial Catalog=EMSDB_QA;User ID=opsdev;Password=opsdev@123;");
        //Production
        SqlConnection con = new SqlConnection("Data Source=BLRPRODRTM\\RTM_PROD_BLR;Initial Catalog=WorkFlowManagerDB;User ID=sa;Password=Prodrtm@123;");
      
        DateTime CurrentDate = DateTime.Now;
        string dtnow;
        DataTable dt = new DataTable();
        string Client;
        string Source;
        string Path;
        string FileName;
        public Form1()
        {
            InitializeComponent();
            GetClients();
            Type();
        }

        public void Type()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Type", typeof(string));           


            DataRow dtrow = dt.NewRow();    // Create New Row //Bind Data to Columns
            dtrow["Type"] = "Digitizing OutLook";
            dt.Rows.Add(dtrow);
            dtrow = dt.NewRow();
            dtrow["Type"] = "PDF Portal";
            dt.Rows.Add(dtrow);
            dtrow = dt.NewRow();
            dtrow["Type"] = "Scanned Mail";       
            dt.Rows.Add(dtrow);

            dtrow = dt.NewRow();
            dtrow["Type"] = "DNR";
            dt.Rows.Add(dtrow);

            //DataRow dr;
            //dr = dt.NewRow();
            //dr.ItemArray = new object[] { 0, "--Select Source--" };
            //dt.Rows.InsertAt(dr, 0);
            ddlSource.ValueMember = "Type";
            ddlSource.DisplayMember = "Type";
            ddlSource.DataSource = dt;             

        }
        private DataTable GetClients()
        {
            DataTable dt = new DataTable();
            using (SqlDataAdapter ada = new SqlDataAdapter("select * from Vodafone_Clients where C_Status=1 and C_Code not in ('3M','CSC','Exxon','GLENCORE','MMC','HELLO WORLD','Honda','HP','TDAmeritrade')", con))
                ada.Fill(dt);

            DataRow dr;  
            dr = dt.NewRow();
            //dr.ItemArray = new object[] { 0, "--Select Client--" };
            //dt.Rows.InsertAt(dr, 0);

            ddlClient.ValueMember = "C_Name";
            ddlClient.DisplayMember = "C_Name";
            ddlClient.DataSource = dt; 
            return dt;
        }

        //Data insertion 
        private void DataInsertion(string Location, string Type)
        {
            dtnow = CurrentDate.ToString("MM dd yyyy");
            string result = GetID();
            int id = Convert.ToInt32(result) + 1;
           
                string foldername = string.Empty;
                string _originalFoldername = string.Empty;
                string filesource = string.Empty;
                string filedest = string.Empty;
                string _originalFiledest = string.Empty;
                string destination = string.Empty;
                string _originalDest = string.Empty;
                string Year = DateTime.Now.Year.ToString();
                string month = DateTime.Now.ToString("MMMM");

                string FileName = Location.Substring(Location.LastIndexOf("\\") + 1);
                string folder = Location.Remove(Location.LastIndexOf("\\") + 1);
                foldername = folder.ToString();

                DirectoryInfo di = new DirectoryInfo(folder);
                if (di.Exists)
                {
                                      
                    Path = foldername;
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
                                    cmd.Parameters.AddWithValue("@FI_ClientCode", Client.ToString());
                                    //cmd.Parameters.AddWithValue("@FI_FileName", Filename);
                                    cmd.Parameters.AddWithValue("@FI_FileName", id + "_" + Client.ToString() + "_" + Filename);
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

              //  }            
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

        private void btnUpload_Click(object sender, EventArgs e)
        {
            string path, Type;

            try
            {

                Client = ddlClient.SelectedValue.ToString();
                Source = ddlSource.SelectedValue.ToString();
         

                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                openFileDialog1.InitialDirectory = @"C:\";

                openFileDialog1.Title = "Browse Text Files";

                openFileDialog1.CheckFileExists = true;

                openFileDialog1.CheckPathExists = true;

                openFileDialog1.DefaultExt = "txt";

                openFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

                openFileDialog1.FilterIndex = 2;

                openFileDialog1.RestoreDirectory = true;

                openFileDialog1.ReadOnlyChecked = true;

                openFileDialog1.ShowReadOnly = true;
                openFileDialog1.Multiselect = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string tempPath = openFileDialog1.FileName.ToString();
               
                        DataInsertion(tempPath, Source);
                }

                MessageBox.Show("Invoice Uploaded Successfully.");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
