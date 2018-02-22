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
using System.Globalization;

namespace TKMOC
{
    public partial class frmMOCCOPMA : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
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

        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string EDITSTATUS;

        public frmMOCCOPMA()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCHMOCCOPMA()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


               
                sbSql.AppendFormat(@" SELECT [ID] AS '代號',[KIND] AS '分類',[NAME] AS '名稱' FROM [TKMOC].[dbo].[MOCCOPMA] ORDER BY [ID] ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

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

        public void SAVEMOCCOPMA()
        {
            if (EDITSTATUS.Equals("ADD"))
            {
                ADDMOCCOPMA();
            }
            else if (EDITSTATUS.Equals("UPDATE"))
            {
                UPDATEMOCCOPMA();
            }
            else
            {
                MessageBox.Show("存檔失敗");
            }
        }

        public void ADDMOCCOPMA()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCCOPMA]");
                sbSql.AppendFormat(" ([ID],[NAME],[KIND])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", textBox1.Text, textBox2.Text,comboBox1.Text);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");


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

        public void UPDATEMOCCOPMA()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCCOPMA]");
                sbSql.AppendFormat(" SET [NAME]='{0}',[KIND]='{1}'", textBox2.Text, comboBox1.Text);
                sbSql.AppendFormat(" WHERE [ID]='{0}'",textBox1.Text);
                sbSql.AppendFormat(" ");


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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["代號"].Value.ToString();
                    textBox2.Text = row.Cells["名稱"].Value.ToString();
                    comboBox1.Text= row.Cells["分類"].Value.ToString();
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                }
            }
        }

        public void DELMOCCOPMA()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE  [TKMOC].[dbo].[MOCCOPMA] WHERE [ID]='{0}'",textBox1.Text);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");


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

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCCOPMA();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            EDITSTATUS = "ADD";
            textBox1.Text = null;
            textBox2.Text = null;

            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            EDITSTATUS = "UPDATE";
        
            //textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            SAVEMOCCOPMA();
            SEARCHMOCCOPMA();
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;
            EDITSTATUS = null;
        }


        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCCOPMA();
                SEARCHMOCCOPMA();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        #endregion


    }
}
