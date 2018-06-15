﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmMOCINVMERGE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string SALSESID = null;
        int result;

        public Report report1 { get; private set; }

        public frmMOCINVMERGE()
        {
            InitializeComponent();

            combobox2load();

            DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
            dgvc.Width = 60;
            dgvc.Name = "選取";

            //新增到DataGridView內的第0欄
            this.dataGridView1.Columns.Insert(0, dgvc);
        }


        #region FUNCTION
        public void combobox2load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT MD002,MD001 FROM [TK].dbo.CMSMD WHERE MD002 LIKE '新%' ORDER BY MD001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD001";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();



        }
        public void SERACH()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TC001 AS '領料單',TC002 AS '單號',TC005  AS '線別' ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTC");
                sbSql.AppendFormat(@"  WHERE TC003>='{0}' AND TC003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TC005='{0}'",comboBox2.SelectedValue.ToString());
                sbSql.AppendFormat(@"  ORDER BY TC001,TC002,TC005");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];

                    

                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ADDDATA()
        {
            DataTable dt = new DataTable();
           
            // Declare DataColumn and DataRow variables.
            DataColumn column;
            DataRow row;
            DataView view;

            // Create new DataColumn, set DataType, ColumnName and add to DataTable.    
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "單別";
            dt.Columns.Add(column);

            // Create second column.
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "單號";
            dt.Columns.Add(column);

            dt.Clear();

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    //MessageBox.Show(dr.Cells[2].Value.ToString());

                    //建立一筆新的DataRow，並且等於新的dt row
                    row = dt.NewRow();

                    //指定每個欄位要儲存的資料
                    row["單別"] = dr.Cells[1].Value.ToString();
                    row["單號"] = dr.Cells[2].Value.ToString();

                    //新增資料至DataTable的dt內
                    dt.Rows.Add(row);
                }
            }


            if (dt.Rows.Count == 0)
            {
                dataGridView2.DataSource = null;
            }
            else if(dt.Rows.Count >=1)
            {
                dataGridView2.DataSource = dt;
            }

           
        }

        public void SETREPORT()
        {
            if(dataGridView2.Rows.Count>=1)
            {
                foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                {
                    MessageBox.Show(dr.Cells[0].Value.ToString()+ dr.Cells[1].Value.ToString());
                   
                }
               
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = null;

            SERACH();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETREPORT();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADDDATA();
        }


        #endregion

       
    }
}
