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
                                    SET [MANU1ACT]=0,[MANU2ACT]=0
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

        #endregion


    }
}
