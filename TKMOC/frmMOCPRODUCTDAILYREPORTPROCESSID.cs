using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;

namespace TKMOC
{
    public partial class frmMOCPRODUCTDAILYREPORTPROCESSID : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        DataGridViewRow drMOCPRODUCTDAILYREPORT = new DataGridViewRow();
        string tablename = null;
        string ID;
        int result;
        Thread TD;
        public frmMOCPRODUCTDAILYREPORTPROCESSID()
        {
            InitializeComponent();
        }
        public frmMOCPRODUCTDAILYREPORTPROCESSID(string SOURCEID)
        {
            InitializeComponent();
            if(!string.IsNullOrEmpty(SOURCEID))
            {
                textBox2.Text = SOURCEID;
                Search(SOURCEID);
            }
        }
        #region FUNCTION
        public void Search(string ID)
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT [PROCESSID] AS '製令',[SOURCEID] AS '來源ID',[ID] AS 'ID' FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORTPROCESSID] WHERE [SOURCEID]='{0}'", ID.ToString());



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //label1.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {

                        //labelSearch.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["製令"].ToString();
                        //textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["來源ID"].ToString();
                        textBox3.Text = ds.Tables["TEMPds1"].Rows[0]["ID"].ToString();
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
        public void UPDATE()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MOCPRODUCTDAILYREPORTPROCESSID] ");
                sbSql.AppendFormat("  SET [PROCESSID]='{1}',[SOURCEID]='{2}' WHERE [ID]='{0}' ", textBox3.Text.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString());
                sbSql.Append("   ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                   

                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public void ADD()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MOCPRODUCTDAILYREPORTPROCESSID] ");
                sbSql.Append(" ([ID],[SOURCEID],[PROCESSID])  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}') ", Guid.NewGuid(),textBox2.Text.ToString(),textBox1.Text.ToString());

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                    

                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SETADD()
        {
            textBox1.ReadOnly = false;
            textBox1.Text = null;
            textBox3.Text = null;
        }

        public void SETUPDATE()
        {
            textBox1.ReadOnly = false;
        }

        public void SETFINISH()
        {
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;
            textBox3.ReadOnly = true;
        }
        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox3.Text.ToString()))
            {
                UPDATE();
            }
            else
            {
                ADD();
            }
            Search(textBox2.Text.ToString());
            SETFINISH();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETADD();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETUPDATE();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion


    }
}
