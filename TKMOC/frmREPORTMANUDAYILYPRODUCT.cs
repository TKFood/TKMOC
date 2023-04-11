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
using Calendar.NET;
using TKITDLL;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmREPORTMANUDAYILYPRODUCT : Form
    {
        public Report report1 { get; private set; }

        public frmREPORTMANUDAYILYPRODUCT()
        {
            InitializeComponent();

            tabControl1.SelectedIndex = 1;
        }

        #region FUNCTION
        public void SEARCH(string MANUDATE)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            DateTime DT = Convert.ToDateTime(MANUDATE);

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
                                    CONVERT(NVARCHAR,[MANUDATE],112) AS '預排日'
                                    ,[MANU1PUR]  AS '小線產能'
                                    ,[MANU1ACT] AS '小線桶數'
                                    ,[MANU2PUR] AS '大線產能'
                                    ,[MANU2ACT] AS '大線桶數'
                                    ,[MANU3PUR] AS '手工產能'
                                    ,[MANU3ACT] AS '手工預排'
                                    ,[MANU4PUR] AS '外包產能'
                                    ,[MANU4ACT] AS '外包預排'
                                    ,(CASE WHEN [MANU1PUR]>0 AND [MANU1ACT]>0 THEN CONVERT(DECIMAL(16,2),([MANU1ACT]/[MANU1PUR])*100) ELSE 0 END ) AS '小線訂單稼動率'
                                    ,(CASE WHEN [MANU2PUR]>0 AND [MANU2ACT]>0 THEN CONVERT(DECIMAL(16,2),([MANU2ACT]/[MANU2PUR])*100) ELSE 0 END ) AS '大線訂單稼動率'
                                    ,(CASE WHEN [MANU3PUR]>0 AND [MANU3ACT]>0 THEN CONVERT(DECIMAL(16,2),([MANU3ACT]/[MANU3PUR])*100) ELSE 0 END ) AS '手工訂單稼動率'
                                    ,(CASE WHEN [MANU4PUR]>0 AND [MANU4ACT]>0 THEN CONVERT(DECIMAL(16,2),([MANU4ACT]/[MANU4PUR])*100) ELSE 0 END ) AS '外包訂單稼動率'
                                    
                                    FROM [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    WHERE CONVERT(NVARCHAR,[MANUDATE],112) LIKE '{0}%'
                                    ORDER BY CONVERT(NVARCHAR,[MANUDATE],112)

                                    ", DT.ToString("yyyyMM"));




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
                sqlConn.Close();
            }
        }
        public void SEARCH2(string MANUDATE)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            DateTime DT = Convert.ToDateTime(MANUDATE);

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
                                    CONVERT(NVARCHAR,[MANUDATE],112) AS '預排日'
                                  
                                    ,[MANU1PURTIMES]  AS '小線產能時數'
                                    ,[MANU1ACTTIMES] AS '小線桶數時數'
                                    ,[MANU2PURTIMES] AS '大線產能時數'
                                    ,[MANU2ACTTIMES] AS '大線桶數時數'
                                    ,[MANU3PURTIMES] AS '手工產能時數'
                                    ,[MANU3ACTTIMES] AS '手工預排時數'
                                    ,[MANU4PURTIMES] AS '外包產能時數'
                                    ,[MANU4ACTTIMES] AS '外包預排時數'
                                    ,(CASE WHEN [MANU1PURTIMES]>0 AND [MANU1ACTTIMES]>0 THEN CONVERT(DECIMAL(16,2),([MANU1ACTTIMES]/[MANU1PURTIMES])*100) ELSE 0 END ) AS '小線訂單稼動率'
                                    ,(CASE WHEN [MANU2PURTIMES]>0 AND [MANU2ACTTIMES]>0 THEN CONVERT(DECIMAL(16,2),([MANU2ACTTIMES]/[MANU2PURTIMES])*100) ELSE 0 END ) AS '大線訂單稼動率'
                                    ,(CASE WHEN [MANU3PURTIMES]>0 AND [MANU3ACTTIMES]>0 THEN CONVERT(DECIMAL(16,2),([MANU3ACTTIMES]/[MANU3PURTIMES])*100) ELSE 0 END ) AS '手工訂單稼動率'
                                    ,(CASE WHEN [MANU4PURTIMES]>0 AND [MANU4ACTTIMES]>0 THEN CONVERT(DECIMAL(16,2),([MANU4ACTTIMES]/[MANU4PURTIMES])*100) ELSE 0 END ) AS '外包訂單稼動率'

                                    
                                    FROM [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    WHERE CONVERT(NVARCHAR,[MANUDATE],112) LIKE '{0}%'
                                    ORDER BY CONVERT(NVARCHAR,[MANUDATE],112)

                                    ", DT.ToString("yyyyMM"));




                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
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

                    textBox1.Text = row.Cells["預排日"].Value.ToString();
                    textBox2.Text = row.Cells["小線產能"].Value.ToString();
                    textBox3.Text = row.Cells["小線桶數"].Value.ToString();
                    textBox4.Text = row.Cells["大線產能"].Value.ToString();
                    textBox5.Text = row.Cells["大線桶數"].Value.ToString();
                    textBox6.Text = row.Cells["手工產能"].Value.ToString();
                    textBox7.Text = row.Cells["手工預排"].Value.ToString();
                    textBox8.Text = row.Cells["外包產能"].Value.ToString();
                    textBox9.Text = row.Cells["外包預排"].Value.ToString();



                }
                else
                {

                }
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    textBox14.Text = row.Cells["預排日"].Value.ToString();
                    textBox15.Text = row.Cells["小線產能時數"].Value.ToString();
                    textBox16.Text = row.Cells["小線桶數時數"].Value.ToString();
                    textBox17.Text = row.Cells["大線產能時數"].Value.ToString();
                    textBox18.Text = row.Cells["大線桶數時數"].Value.ToString();
                    textBox19.Text = row.Cells["手工產能時數"].Value.ToString();
                    textBox20.Text = row.Cells["手工預排時數"].Value.ToString();
                    textBox21.Text = row.Cells["外包產能時數"].Value.ToString();
                    textBox22.Text = row.Cells["外包預排時數"].Value.ToString();



                }
                else
                {

                }
            }
        }
        public void UPDATE_DATEILS(string MANUDATE,string MANU1PUR, string MANU1ACT, string MANU2PUR, string MANU2ACT, string MANU3PUR, string MANU3ACT, string MANU4PUR, string MANU4ACT)
        {
            SqlConnection sqlConn = new SqlConnection();
            StringBuilder sbSql = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;

            if (!string.IsNullOrEmpty(MANUDATE))
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

                                   UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU1PUR]='{1}'
                                    ,[MANU1ACT]='{2}'
                                    ,[MANU2PUR]='{3}'
                                    ,[MANU2ACT]='{4}'
                                    ,[MANU3PUR]='{5}'
                                    ,[MANU3ACT]='{6}'
                                    ,[MANU4PUR]='{7}'
                                    ,[MANU4ACT]='{8}'
                                    WHERE  CONVERT(NVARCHAR,[MANUDATE],112)='{0}'
                                   
                                    ", MANUDATE, MANU1PUR, MANU1ACT, MANU2PUR, MANU2ACT, MANU3PUR, MANU3ACT, MANU4PUR, MANU4ACT);


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
        }

        public void UPDATE_DATEILS2(string MANUDATE, string MANU1PURTIMES, string MANU1ACTTIMES, string MANU2PURTIMES, string MANU2ACTTIMES, string MANU3PURTIMES, string MANU3ACTTIMES, string MANU4PURTIMES, string MANU4ACTTIMES)
        {
            SqlConnection sqlConn = new SqlConnection();
            StringBuilder sbSql = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;

            if (!string.IsNullOrEmpty(MANUDATE))
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

                                   UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU1PUR]='{1}'
                                    ,[MANU1ACTTIMES]='{2}'
                                    ,[MANU2PURTIMES]='{3}'
                                    ,[MANU2ACTTIMES]='{4}'
                                    ,[MANU3PURTIMES]='{5}'
                                    ,[MANU3ACTTIMES]='{6}'
                                    ,[MANU4PURTIMES]='{7}'
                                    ,[MANU4ACTTIMES]='{8}'
                                    WHERE  CONVERT(NVARCHAR,[MANUDATE],112)='{0}'
                                   
                                    ", MANUDATE, MANU1PURTIMES, MANU1ACTTIMES, MANU2PURTIMES, MANU2ACTTIMES, MANU3PURTIMES, MANU3ACTTIMES, MANU4PURTIMES, MANU4ACTTIMES);


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
        }

        public void UPDATE_PUR(string SDATES, string EDATES, string MANU1PUR, string MANU2PUR, string MANU3PUR, string MANU4PUR)
        {

            SqlConnection sqlConn = new SqlConnection();
            StringBuilder sbSql = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;

            if (!string.IsNullOrEmpty(SDATES))
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
                                    UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU1PUR]='{2}'
                                    ,[MANU2PUR]='{3}'
                                    ,[MANU3PUR]='{4}'
                                    ,[MANU4PUR]='{5}'
                                    WHERE  CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'
                                  
                                   
                                    ",  SDATES,  EDATES,  MANU1PUR,  MANU2PUR,  MANU3PUR,  MANU4PUR);


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

        }

        public void UPDATE_PUR2(string SDATES, string EDATES, string MANU1PURTIMES, string MANU2PURTIMES, string MANU3PURTIMES, string MANU4PURTIMES)
        {

            SqlConnection sqlConn = new SqlConnection();
            StringBuilder sbSql = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;

            if (!string.IsNullOrEmpty(SDATES))
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
                                    UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU1PURTIMES]='{2}'
                                    ,[MANU2PURTIMES]='{3}'
                                    ,[MANU3PURTIMES]='{4}'
                                    ,[MANU4PURTIMES]='{5}'
                                    WHERE  CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'
                                  
                                   
                                    ", SDATES, EDATES, MANU1PURTIMES, MANU2PURTIMES, MANU3PURTIMES, MANU4PURTIMES);


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

        }

        public void UPDATE_ACT(string SDATES, string EDATES)
        {
            SqlConnection sqlConn = new SqlConnection();
            StringBuilder sbSql = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;

            if (!string.IsNullOrEmpty(SDATES) && !string.IsNullOrEmpty(EDATES))
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
                                    UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU1ACT]=0,[MANU2ACT]=0,[MANU3ACT]=0
                                    WHERE  CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'

                                    UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU1ACT]=TEMP.SUMBAR
                                    FROM (
                                    SELECT 
                                    [MANUDATE]      
                                    ,[MANU]     
                                    ,ISNULL(SUM([BAR]),0) AS 'SUMBAR'
                                    FROM [TKMOC].[dbo].[MOCMANULINE]
                                    WHERE [MANU]  IN ('製一線')
                                    AND [MB001] NOT IN (SELECT  [MB001]   FROM [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                                    AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'
                                    GROUP BY [MANUDATE],[MANU]     
                                    ) AS TEMP
                                    WHERE TEMP.MANUDATE=[MANUDAYILYPRODUCT].[MANUDATE]
                                 

                                    UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU2ACT]=TEMP.SUMBAR
                                    FROM (
                                    SELECT 
                                    [MANUDATE]      
                                    ,[MANU]     
                                    ,ISNULL(SUM([BAR]),0) AS 'SUMBAR'
                                    FROM [TKMOC].[dbo].[MOCMANULINE]
                                    WHERE [MANU]  IN ('製二線')
                                    AND [MB001] NOT IN (SELECT  [MB001]   FROM [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                                    AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'
                                    GROUP BY [MANUDATE],[MANU]     
                                    ) AS TEMP
                                    WHERE TEMP.MANUDATE=[MANUDAYILYPRODUCT].[MANUDATE]


                                    UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU3ACT]=TEMP.SUMPACKAGE
                                    FROM (
                                    SELECT 
                                    [MANUDATE]      
                                    ,[MANU]     
                                    ,ISNULL(SUM([PACKAGE]),0) AS 'SUMPACKAGE'
                                    FROM [TKMOC].[dbo].[MOCMANULINE]
                                    WHERE [MANU]  IN ('包裝線')
                                    AND [MB001] NOT IN (SELECT  [MB001]   FROM [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                                    AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'
                                    GROUP BY [MANUDATE],[MANU]     
                                    ) AS TEMP
                                    WHERE TEMP.MANUDATE=[MANUDAYILYPRODUCT].[MANUDATE]

                                    ", SDATES, EDATES);


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
        }

        public void UPDATE_ACT2(string SDATES, string EDATES)
        {
            int PACKAGEDAYSHR = 8;
            SqlConnection sqlConn = new SqlConnection();
            StringBuilder sbSql = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;

            if (!string.IsNullOrEmpty(SDATES) && !string.IsNullOrEmpty(EDATES))
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
                                    UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU1ACTTIMES]=0,[MANU2ACTTIMES]=0,[MANU3ACTTIMES]=0
                                    WHERE  CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'

                                    UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU1ACTTIMES]=TEMP.SUMBAR
                                    FROM (
                                        SELECT 
                                        [MANUDATE]      
                                        ,[MANU]     
                                        ,ISNULL(SUM([BAR]*[ERPINVMB].[BUCKETTIMES]),0) AS 'SUMBAR'
                                        FROM [TKMOC].[dbo].[MOCMANULINE]
                                        LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [MOCMANULINE].[MB001]=[ERPINVMB].[MB001]
                                        WHERE [MANU]  IN ('製一線')
                                        AND [MOCMANULINE].[MB001] NOT IN (SELECT  [MB001]   FROM [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                                        AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'
                                        GROUP BY [MANUDATE],[MANU]   
                                    ) AS TEMP
                                    WHERE TEMP.MANUDATE=[MANUDAYILYPRODUCT].[MANUDATE]
                                 

                                    UPDATE [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    SET [MANU3ACTTIMES]=TEMP.SUMPACKAGE
                                    FROM (
                                        SELECT 
                                        [MANUDATE]      
                                        ,[MANU]     
                                        ,ISNULL(SUM([PACKAGE]/[ERPINVMB].[PACKAGETIMES]*{2}),0) AS 'SUMPACKAGE'
                                        FROM [TKMOC].[dbo].[MOCMANULINE]
                                        LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [MOCMANULINE].[MB001]=[ERPINVMB].[MB001]
                                        WHERE [MANU]  IN ('包裝線')
                                        AND [MOCMANULINE].[MB001] NOT IN (SELECT  [MB001]   FROM [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                                         AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'
                                        GROUP BY [MANUDATE],[MANU]     
                                    ) AS TEMP
                                    WHERE TEMP.MANUDATE=[MANUDAYILYPRODUCT].[MANUDATE]

                                    ", SDATES, EDATES, PACKAGEDAYSHR);


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
        }

        public void SETFASTREPORT(string MANUDATE)
        {
            SqlConnection sqlConn = new SqlConnection();
            StringBuilder sbSql = new StringBuilder();

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\稼動率.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL(MANUDATE);
            Table.SelectCommand = SQL;

            //report1.SetParameterValue("P1", COMMENT);
            //report1.SetParameterValue("P2", NUM);

            report1.Preview = previewControl1;
            report1.Show();
        }


        public string SETFASETSQL(string MANUDATE)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();
                       
            FASTSQL.AppendFormat(@"    
                                SELECT 
                                CONVERT(NVARCHAR,[MANUDATE],112) AS '預排日'
                                ,[MANU1PUR]  AS '小線產能'
                                ,[MANU1ACT] AS '小線桶數'
                                ,[MANU2PUR] AS '大線產能'
                                ,[MANU2ACT] AS '大線桶數'
                                ,[MANU3PUR] AS '手工產能'
                                ,[MANU3ACT] AS '手工預排'
                                ,[MANU4PUR] AS '外包產能'
                                ,[MANU4ACT] AS '外包預排'
                                ,(CASE WHEN [MANU1PUR]>0 AND [MANU1ACT]>0 THEN CONVERT(DECIMAL(16,2),([MANU1ACT]/[MANU1PUR])*100) ELSE 0 END ) AS '小線訂單稼動率'
                                ,(CASE WHEN [MANU2PUR]>0 AND [MANU2ACT]>0 THEN CONVERT(DECIMAL(16,2),([MANU2ACT]/[MANU2PUR])*100) ELSE 0 END ) AS '大線訂單稼動率'
                                ,(CASE WHEN [MANU3PUR]>0 AND [MANU3ACT]>0 THEN CONVERT(DECIMAL(16,2),([MANU3ACT]/[MANU3PUR])*100) ELSE 0 END ) AS '手工訂單稼動率'
                                ,(CASE WHEN [MANU4PUR]>0 AND [MANU4ACT]>0 THEN CONVERT(DECIMAL(16,2),([MANU4ACT]/[MANU4PUR])*100) ELSE 0 END ) AS '外包訂單稼動率'
                                FROM [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                WHERE CONVERT(NVARCHAR,[MANUDATE],112) LIKE '{0}%'
                                ORDER BY CONVERT(NVARCHAR,[MANUDATE],112)

                                    ", MANUDATE);

            return FASTSQL.ToString(); 
        }

        public void SETFASTREPORT2(string MANUDATE)
        {
            SqlConnection sqlConn = new SqlConnection();
            StringBuilder sbSql = new StringBuilder();

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\稼動率V2.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL2(MANUDATE);
            Table.SelectCommand = SQL;

            //report1.SetParameterValue("P1", COMMENT);
            //report1.SetParameterValue("P2", NUM);

            report1.Preview = previewControl2;
            report1.Show();
        }


        public string SETFASETSQL2(string MANUDATE)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"    
                                SELECT 
                                CONVERT(NVARCHAR,[MANUDATE],112) AS '預排日'
                                ,[MANU1PURTIMES] AS '小線產能時數'
                                ,[MANU1ACTTIMES] AS '小線桶數時數'
                                ,[MANU2PURTIMES] AS '大線產能時數'
                                ,[MANU2ACTTIMES] AS '大線桶數時數'
                                ,[MANU3PURTIMES] AS '手工產能時數'
                                ,[MANU3ACTTIMES] AS '手工預排時數'
                                ,[MANU4PURTIMES] AS '外包產能時數'
                                ,[MANU4ACTTIMES] AS '外包預排時數'
                                ,(CASE WHEN [MANU1PURTIMES]>0 AND [MANU1ACTTIMES]>0 THEN CONVERT(DECIMAL(16,2),([MANU1ACTTIMES]/[MANU1PURTIMES])*100) ELSE 0 END ) AS '小線訂單稼動率'
                                ,(CASE WHEN [MANU2PURTIMES]>0 AND [MANU2ACTTIMES]>0 THEN CONVERT(DECIMAL(16,2),([MANU2ACTTIMES]/[MANU2PURTIMES])*100) ELSE 0 END ) AS '大線訂單稼動率'
                                ,(CASE WHEN [MANU3PURTIMES]>0 AND [MANU3ACTTIMES]>0 THEN CONVERT(DECIMAL(16,2),([MANU3ACTTIMES]/[MANU3PURTIMES])*100) ELSE 0 END ) AS '手工訂單稼動率'
                                ,(CASE WHEN [MANU4PURTIMES]>0 AND [MANU4ACTTIMES]>0 THEN CONVERT(DECIMAL(16,2),([MANU4ACTTIMES]/[MANU4PURTIMES])*100) ELSE 0 END ) AS '外包訂單稼動率'
                                FROM [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                WHERE CONVERT(NVARCHAR,[MANUDATE],112) LIKE '{0}%'
                                ORDER BY CONVERT(NVARCHAR,[MANUDATE],112)

                                    ", MANUDATE);

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            UPDATE_DATEILS(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text);
            SEARCH(dateTimePicker1.Value.ToString("yyyy/MM/dd"));
        }
       
        private void button2_Click(object sender, EventArgs e)
        {
            UPDATE_PUR(dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"), textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text);
            SEARCH(dateTimePicker1.Value.ToString("yyyy/MM/dd"));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyy/MM/dd"));
        }

        private void button4_Click(object sender, EventArgs e)
        {
            UPDATE_ACT(dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
            SEARCH(dateTimePicker1.Value.ToString("yyyy/MM/dd"));
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker6.Value.ToString("yyyyMM"));
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SEARCH2(dateTimePicker7.Value.ToString("yyyy/MM/dd"));
        }
        private void button7_Click(object sender, EventArgs e)
        {
            UPDATE_DATEILS2(textBox14.Text, textBox15.Text, textBox16.Text, textBox17.Text, textBox18.Text, textBox19.Text, textBox20.Text, textBox21.Text, textBox22.Text);
            SEARCH2(dateTimePicker7.Value.ToString("yyyy/MM/dd"));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            UPDATE_PUR2(dateTimePicker8.Value.ToString("yyyyMMdd"), dateTimePicker9.Value.ToString("yyyyMMdd"), textBox23.Text, textBox24.Text, textBox25.Text, textBox26.Text);
            SEARCH2(dateTimePicker7.Value.ToString("yyyy/MM/dd"));
        }
        private void button9_Click(object sender, EventArgs e)
        {
            UPDATE_ACT2(dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"));
            SEARCH2(dateTimePicker7.Value.ToString("yyyy/MM/dd"));
        }

        private void button10_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker12.Value.ToString("yyyyMM"));
        }
        #endregion


    }
}
