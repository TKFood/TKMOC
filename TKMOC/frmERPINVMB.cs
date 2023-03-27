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
using TKITDLL;

namespace TKMOC
{
    public partial class frmERPINVMB : Form
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
        int rownum = 0;
        Thread TD;

        public frmERPINVMB()
        {
            InitializeComponent();
        }
        #region FUNCTION
        public void Search()
        {
            StringBuilder Query = new StringBuilder();
            if(!string.IsNullOrEmpty(textBox4.Text.ToString()))
            {
                Query.AppendFormat(@" AND MB001 LIKE '{0}%' ", textBox4.Text.ToString());
            }

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

           
                sbSql.AppendFormat(@" 
                                    SELECT [MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格' ,[PROCESSNUM] AS '標準批量',[PROCESSTIME] AS '標準時間'
                                    ,[BUCKETTIMES] AS '1桶的生產時間'
                                    FROM [TKMOC].[dbo].[ERPINVMB] 
                                    WHERE 1=1 
                                    {0}
                                    ORDER BY [MB001]
 
                                    ", Query.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    labelSearch.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {

                        labelSearch.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.Columns["品號"].Width = 200; // 設定指定欄位的寬度為 100 像素
                        dataGridView1.Columns["品名"].Width = 100; // 設定指定欄位的寬度為 100 像素
                        dataGridView1.Columns["規格"].Width = 100; // 設定指定欄位的寬度為 100 像素
                        dataGridView1.Columns["標準批量"].Width = 100; // 設定指定欄位的寬度為 100 像素
                        dataGridView1.Columns["標準時間"].Width = 100; // 設定指定欄位的寬度為 100 像素
                        dataGridView1.Columns["1桶的生產時間"].Width = 100; // 設定指定欄位的寬度為 100 像素

                        dataGridView1.CurrentCell = dataGridView1[0, rownum];
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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                
                sbSql.AppendFormat(@"  
                                    UPDATE [TKMOC].[dbo].[ERPINVMB] 
                                    SET [MB002]='{1}',[MB003]='{2}',[PROCESSNUM]='{3}',[PROCESSTIME]='{4}' ,[BUCKETTIMES]='{5}'
                                    WHERE [MB001]='{0}' 
                                    ", textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString(), textBox5.Text.ToString());

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
                    rownum = dataGridView1.CurrentCell.RowIndex;

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

        public void SetUPDATE()
        {
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
        }

        public void SetFINISH()
        {
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;
            textBox3.ReadOnly = true;
        }

        public void ADDNEW()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append("INSERT INTO [TKMOC].dbo.ERPINVMB (MB001,MB002,MB003,[PROCESSNUM] ,[PROCESSTIME]) ");
                sbSql.AppendFormat(" SELECT REPLACE(MB001,' ','') MB001,MB002,MB003,0,0 FROM [TK].dbo.INVMB WITH (NOLOCK) WHERE (MB001 LIKE '4%' OR MB001 LIKE '3%' ) AND MB001 NOT IN (SELECT  REPLACE(MB001,' ','') FROM [TKMOC].dbo.ERPINVMB WITH (NOLOCK) )");
                sbSql.Append("  ");

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
            if (dataGridView1.Rows.Count >= 1)
            {
                drMOCPRODUCTDAILYREPORT = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];
                textBox1.Text = drMOCPRODUCTDAILYREPORT.Cells["品號"].Value.ToString();
                textBox2.Text = drMOCPRODUCTDAILYREPORT.Cells["品名"].Value.ToString();
                numericUpDown1.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["標準批量"].Value.ToString());
                numericUpDown2.Value = Convert.ToDecimal(drMOCPRODUCTDAILYREPORT.Cells["標準時間"].Value.ToString());
                textBox5.Text = drMOCPRODUCTDAILYREPORT.Cells["1桶的生產時間"].Value.ToString();


            }
        }

        public void UPDATEERPINVMB()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  UPDATE [TKMOC].dbo.ERPINVMB");
                sbSql.AppendFormat("  SET ERPINVMB.MB002=INVMB.MB002,ERPINVMB.MB003=INVMB.MB003");
                sbSql.AppendFormat("  FROM [TK].dbo.INVMB");
                sbSql.AppendFormat("  WHERE ERPINVMB.MB001=INVMB.MB001");
                sbSql.AppendFormat("  AND (ERPINVMB.MB002<>INVMB.MB002 OR ERPINVMB.MB003<>INVMB.MB003)");
                sbSql.AppendFormat("  ");

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
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADDNEW();
            UPDATEERPINVMB();
            Search();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SetUPDATE();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            UPDATE();
            SetFINISH();
            Search();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            UPDATEERPINVMB();
            Search();
        }
        #endregion


    }
}
