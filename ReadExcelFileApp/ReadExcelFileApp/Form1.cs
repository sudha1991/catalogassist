using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelFileApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Visible = false;
        }

        private void btnChoose_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog();//open dialog to choose file
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)//if there is a file choosen by the user
            {
                filePath = file.FileName;//get the path of the file
                fileExt = Path.GetExtension(filePath);//get the file extension
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath, fileExt);//read excel file
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);//custom messageBox to show error
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();//to close the window(Form1)
        }

        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";//for above excel 2007
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    con.Open();
                    DataTable dt = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                    string SqlQuery = "";
                    bool proceed = false;
                    foreach( DataRow df in dt.Rows){
                        if (df["TABLE_NAME"].ToString() != "filter$")
                        {
                            SqlQuery += (SqlQuery == "" ? "" : " union ") + " select * from [" + df["TABLE_NAME"].ToString()+"] ";
                        }
                        else {
                            proceed = true;
                        }
                    }
                    con.Close();

                    if (proceed)
                    {
                        //SELECT `PRIMARY_KEY`, rand() FROM table ORDER BY rand() LIMIT 5000;
                        DataTable filter = new DataTable();
                        // Read the filter table and making the condition
                        OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [filter$]", con);
                        oleAdpt.Fill(filter);

                        string condition = "";
                        foreach (DataColumn col in filter.Columns)
                        {
                            //condition += (condition == "" ? " WHERE " : " AND ") + " [" + col.ColumnName + "] not IN (select [" + col.ColumnName + "] from [filter$] where [" + col.ColumnName + "] <> '' )";
                            foreach (DataRow row in filter.Rows)
                            {
                                string s = row[col.ColumnName].ToString();
                                if (s.Trim() != "")
                                {
                                    condition += (condition == "" ? "" : " AND ") + " UCase([" + col.ColumnName + "]) not like '%" + s.ToUpper() + "%'";
                                }
                            }
                        }
                        //con.BeginTransaction();
                        //OleDbCommand cmd = new OleDbCommand(SqlQuery + condition, con);
                        //using (OleDbCommand cmd = new OleDbCommand(SqlQuery + condition, con))
                        //{
                        //    //cmd.CommandTimeout = 3600;
                        //    OleDbDataReader rdr = cmd.ExecuteReader();
                        //    dtexcel.Load(rdr);
                        //}
                        //con.Dispose(); - SqlQuery + " WHERE "+condition

                        //var selectStatement = string.Format(SqlQuery + condition);
                        //using (OleDbDataAdapter cmd = new OleDbDataAdapter(SqlQuery + " WHERE "+condition, con))
                        //{
                        //    cmd.Fill(dtexcel);
                        //}


                        System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(SqlQuery + " WHERE " + condition, con);
                        cmd.CommandType = CommandType.Text;
                        new System.Data.OleDb.OleDbDataAdapter(cmd).Fill(dtexcel);

                        //OleDbDataAdapter oleAdptx = new OleDbDataAdapter("select * from (" + SqlQuery + ") x " + condition, con);//here we read data from sheet1
                        //OleDbDataAdapter oleAdptx = new OleDbDataAdapter(SqlQuery + condition, con);//here we read data from sheet1
                        //oleAdptx.Fill(dtexcel);//fill excel data into dataTable
                    }
                    else {
                        throw new Exception("Filter Sheet is missing");
                    }
                }
                catch (Exception ex)
                {
                    if(con.State == ConnectionState.Open) con.Close();
                    MessageBox.Show(ex.Message.ToString());
                }
                finally {
                    
                }
            }
            return dtexcel;
        }
    }
}
