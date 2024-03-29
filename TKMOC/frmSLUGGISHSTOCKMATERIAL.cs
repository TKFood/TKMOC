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
using Calendar.NET;
using Excel = Microsoft.Office.Interop.Excel;
using FastReport;
using FastReport.Data;
using TKITDLL;


namespace TKMOC
{
    public partial class frmSLUGGISHSTOCKMATERIAL : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();

        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        int result;

        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        string tablename = null;


        public frmSLUGGISHSTOCKMATERIAL()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCH(string SDay)
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


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 品號, 品名, 批號, 庫存量, 單位, 在倉日期, 有效天數
                                    ,(SELECT TOP 1 [COMMENTS] FROM  [TKMOC].[dbo].[SLUGGISHSTOCK] WHERE [MB001]=品號 AND [LOTNO]=批號 ORDER BY [ID] DESC)     AS '記錄'       
                                    FROM(
                                    SELECT   LA001 AS '品號', INVMB.MB002 AS '品名', INVMB.MB003 AS '規格', LA016 AS '批號'
                                    , CONVERT(DECIMAL(16, 3), SUM(LA005 * LA011)) AS '庫存量', INVMB.MB004 AS '單位'
                                    , DATEDIFF(DAY, LA016, '{0}') AS '在倉日期old'
                                    , (CASE WHEN DATEDIFF(DAY, LA016, '{0}') >= 0 THEN DATEDIFF(DAY, LA016, '{0}') ELSE(CASE WHEN DATEDIFF(DAY, LA016, '{0}') < 0 THEN (CASE WHEN MB198 = '2' THEN DATEDIFF(DAY, DATEADD(month, -1 * MB023, LA016), '{0}') ELSE (CASE WHEN MB198 = '1' THEN DATEDIFF(DAY, DATEADD(DAY, -1 * MB023, LA016), '{0}') ELSE (CASE WHEN MB198 = '3' THEN DATEDIFF(DAY, DATEADD(YEAR, -1 * MB023, LA016), '{0}') ELSE (0) END) END) END) END) END) AS '在倉日期'
                                    ,(DATEDIFF(DAY, '{0}', LA016))   AS '有效天數'
                                    ,(SELECT TOP 1 TC006 + ' ' + MV002 FROM[TK].dbo.COPTC,[TK].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN(SELECT TOP 1 TA026+TA027 FROM [TK].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG WHERE TG004= LA001 AND TG017 = LA016))) AS '業務'
                                    FROM[TK].dbo.INVLA WITH(NOLOCK)
                                    LEFT JOIN[TK].dbo.INVMB WITH(NOLOCK) ON MB001 = LA001
                                    WHERE(LA009= '20006')
                                    AND LA001 NOT LIKE '122221001'
                                    AND LA001 NOT LIKE '114141009'
                                    AND LA001 NOT LIKE '301%'

                                    GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004
                                    HAVING SUM(LA005* LA011)<>0 
                                    ) AS TEMP
                                    ORDER BY 有效天數 

                                    ", SDay);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["ds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        ////根据列表中数据不同，显示不同颜色背景
                        //foreach (DataGridViewRow dgRow in dataGridView1.Rows)
                        //{
                        //    ////判断
                        //    //if (Convert.ToDecimal(dgRow.Cells[5].Value) > 0)
                        //    //{
                        //    //    //将这行的背景色设置成Pink
                        //    //    dgRow.DefaultCellStyle.BackColor = Color.Pink;

                        //    //}
                        //}

                        dataGridView1.Columns["品號"].Width = 100;
                        dataGridView1.Columns["品名"].Width = 220;



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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox2.Text = row.Cells["品名"].Value.ToString().Trim();
                    textBox3.Text = row.Cells["批號"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["庫存量"].Value.ToString().Trim();
                    textBox6.Text = row.Cells["在倉日期"].Value.ToString().Trim();

                    SEARCHSLUGGISHSTOCKMATERIAL(row.Cells["品號"].Value.ToString().Trim(), row.Cells["批號"].Value.ToString().Trim());
                }
                else
                {

                }
            }
        }

        public void SEARCHSLUGGISHSTOCKMATERIAL(string MB001, string LOTNO)
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


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT
                                    [LOTNO] AS '批號',[NUMS] AS '庫存量',[COMMENTS] AS '記錄',[ID] AS 'ID',[MB001] AS '品號',[MB002] AS '品名'
                                    FROM [TKMOC].[dbo].[SLUGGISHSTOCKMATERIAL]
                                    WHERE [MB001]='{0}' AND [LOTNO]='{1}'
                                    ORDER BY [ID] DESC

                                    ", MB001, LOTNO);

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();


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

        public void ADDSLUGGISHSTOCKMATERIAL(string ID, string MB001, string MB002, string LOTNO, string NUMS, string STAYDAYS, string COMMENTS)
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
                                    INSERT INTO [TKMOC].[dbo].[SLUGGISHSTOCKMATERIAL]
                                    ([ID],[MB001],[MB002],[LOTNO],[NUMS],[STAYDAYS],[COMMENTS])
                                    VALUES
                                    ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')
                                    ", ID, MB001, MB002, LOTNO, NUMS, STAYDAYS, COMMENTS);


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

        public void SETNULL()
        {
            textBox5.Text = null;
        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1(DateTime.Now.ToString("yyyyMMdd"));

            Report report1 = new Report();
            report1.Load(@"REPORT\呆滯表記錄-原料.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string SDay)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"
                                SELECT 品號, 品名, 批號, 庫存量, 單位, 在倉日期, 有效天數
                                ,(SELECT TOP 1 [COMMENTS] FROM  [TKMOC].[dbo].[SLUGGISHSTOCKMATERIAL] WHERE [MB001]=品號 AND [LOTNO]=批號 ORDER BY [ID] DESC)     AS '記錄'       
                                FROM(
                                SELECT   LA001 AS '品號', INVMB.MB002 AS '品名', INVMB.MB003 AS '規格', LA016 AS '批號'
                                , CONVERT(DECIMAL(16, 3), SUM(LA005 * LA011)) AS '庫存量', INVMB.MB004 AS '單位'
                                , DATEDIFF(DAY, LA016, '{0}') AS '在倉日期old'
                                , (CASE WHEN DATEDIFF(DAY, LA016, '{0}') >= 0 THEN DATEDIFF(DAY, LA016, '{0}') ELSE(CASE WHEN DATEDIFF(DAY, LA016, '{0}') < 0 THEN(CASE WHEN MB198 = '2' THEN DATEDIFF(DAY, DATEADD(month, -1 * MB023, LA016), '{0}') END) END) END) AS '在倉日期'
                                 ,(DATEDIFF(DAY, '{0}', LA016))   AS '有效天數'
                                ,(SELECT TOP 1 TC006 + ' ' + MV002 FROM[TK].dbo.COPTC,[TK].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN(SELECT TOP 1 TA026+TA027 FROM [TK].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG WHERE TG004= LA001 AND TG017 = LA016))) AS '業務'
                                FROM[TK].dbo.INVLA WITH(NOLOCK)
                                LEFT JOIN[TK].dbo.INVMB WITH(NOLOCK) ON MB001 = LA001
                                WHERE(LA009= '20006')
                                AND LA001 NOT LIKE '122221001'
                                AND LA001 NOT LIKE '114141009'
                                AND LA001 NOT LIKE '301%'
                                GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004
                                HAVING SUM(LA005* LA011)<>0 
                                ) AS TEMP
                                ORDER BY 有效天數    

                                ", SDay);

            return SB;

        }


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(DateTime.Now.ToString("yyyyMMdd"));
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ADDSLUGGISHSTOCKMATERIAL(DateTime.Now.ToString("yyyyMMddHHmmss"), textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox6.Text, textBox5.Text);

            SEARCHSLUGGISHSTOCKMATERIAL(textBox1.Text, textBox3.Text);
            SETNULL(); 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion


    }
}
