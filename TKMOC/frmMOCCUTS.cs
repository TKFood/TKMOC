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
    public partial class frmMOCCUTS : Form
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

        public frmMOCCUTS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search(string MB001)
        {
            StringBuilder Query = new StringBuilder();
            if (!string.IsNullOrEmpty(MB001))
            {
                Query.AppendFormat(@" AND ([MB001] LIKE '%{0}%' OR [MB002] LIKE '%{0}%') ", MB001);
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
                                    SELECT 
                                    [MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    ,[MB003] AS '規格'
                                    ,[MANULINES] AS '線別'
                                    ,[CUTS] AS '刀模'
                                    ,[WEIGHTS] AS '淨重'
                                    FROM [TKMOC].[dbo].[REPORTCUTS]
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
                        dataGridView1.Columns["線別"].Width = 100; // 設定指定欄位的寬度為 100 像素
                        dataGridView1.Columns["刀模"].Width = 100; // 設定指定欄位的寬度為 100 像素
                        dataGridView1.Columns["淨重"].Width = 100; // 設定指定欄位的寬度為 100 像素

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

            
                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKMOC].[dbo].[REPORTCUTS]
                                    ( MB001,MB002,MB003,[MANULINES])
                                    SELECT MB001,MB002,MB003,MD002
                                    FROM [TK].dbo.INVMB,[TK].dbo.CMSMD
                                    WHERE MB068=MD001
                                    AND (MB001 LIKE '3%' OR  MB001 LIKE '4%')
                                    AND REPLACE(MB001+MD002,' ','') NOT IN (SELECT   REPLACE(MB001+MANULINES,' ','') FROM  [TKMOC].[dbo].[REPORTCUTS])
                                    ");
            

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

                    MessageBox.Show("完成"); 
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
            Search(textBox4 .Text);
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            ADDNEW();
        }
    }
}
