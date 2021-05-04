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
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmREPORTMOCTAB : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        int rownum = 0;
        DataSet ds1 = new DataSet();

        string STATUS = "";

        public Report report1 { get; private set; }

        public frmREPORTMOCTAB()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void frmREPORTMOCTAB_Load(object sender, EventArgs e)
        {

            SETCODE();

            //textBox3.Text = SEARCHMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"));

            //if(!string.IsNullOrEmpty(textBox3.Text))
            //{
            //    SETCODE();
            //}
        }
        public void Search()
        {         
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            StringBuilder Query = new StringBuilder();

            if (!string.IsNullOrEmpty(textBox5.Text.ToString()))
            {
                Query.AppendFormat(@" AND ( MB001 LIKE '{0}%'  OR MB002 LIKE '%{0}%')", textBox5.Text.ToString().Trim());
            }

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT [MB001] AS '品號',[MB002]  AS '品名',[MB003]  AS '規格',[PCT]  AS '比例',[ALLERGEN]  AS '過敏原',[SPEC] AS '餅體'");
                sbSql.AppendFormat(@" FROM [TKMOC].[dbo].[ERPINVMB] ");
                sbSql.AppendFormat(@" WHERE 1=1 ");
                sbSql.AppendFormat(@" {0}", Query.ToString());
                sbSql.AppendFormat(@"  ORDER BY [MB001]");
                sbSql.AppendFormat(@" ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds1.Tables["ds1"];
                        dataGridView2.AutoResizeColumns();

                        dataGridView2.Columns["品號"].Width = 160;
                        dataGridView2.Columns["品名"].Width = 260;
                        dataGridView2.Columns["規格"].Width = 100;
                        dataGridView2.Columns["比例"].Width = 100;
                        dataGridView2.Columns["過敏原"].Width = 100;
                        dataGridView2.Columns["餅體"].Width = 100;

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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count >= 1)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    textBox6.Text = row.Cells["品號"].Value.ToString();
                    textBox7.Text = row.Cells["品名"].Value.ToString();
                    textBox8.Text = row.Cells["規格"].Value.ToString();
                    textBox1.Text = row.Cells["比例"].Value.ToString();
                    textBox2.Text = row.Cells["過敏原"].Value.ToString();
                    textBox4.Text = row.Cells["餅體"].Value.ToString();

                }
                else
                {
                    textBox6.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox4.Text = null;

                }
            }
        }

        public void ADDERPINVMB()
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
                sbSql.AppendFormat(@"INSERT INTO [TKMOC].dbo.ERPINVMB (MB001,MB002,MB003,[PROCESSNUM] ,[PROCESSTIME],[BOXNUM],[BOARDNUM],[PCT],[ALLERGEN]) ");
                sbSql.AppendFormat(@" SELECT MB001,MB002,MB003,0,0,0,0,0,0 FROM [TK].dbo.INVMB WITH (NOLOCK) WHERE (MB001 LIKE '4%' OR MB001 LIKE '3%' ) AND MB001 NOT IN (SELECT MB001 FROM [TKMOC].dbo.ERPINVMB WITH (NOLOCK) )");
                sbSql.AppendFormat(@"  ");

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

        public void UPDATEERPINVMB()
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
                sbSql.AppendFormat(@" UPDATE [TKMOC].[dbo].[ERPINVMB] ");
                sbSql.AppendFormat(@" SET [PCT]='{0}',[ALLERGEN]='{1}',[SPEC]='{3}' WHERE [MB001]='{2}' ", textBox1.Text.ToString().Trim(), textBox2.Text.ToString().Trim(), textBox6.Text.ToString().Trim(), textBox4.Text.ToString().Trim());
                sbSql.AppendFormat(@"  ");

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



        public void SETFASTREPORT(string SDAY, string CODE)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\製令明細表2020.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL(SDAY);
            Table.SelectCommand = SQL;

            report1.SetParameterValue("P1", CODE);
           

            report1.Preview = previewControl1;
            report1.Show();
        }


        public string SETFASETSQL(string SDAY)
        {
            StringBuilder FASTSQL = new StringBuilder();

            //,CASE WHEN TA006 NOT LIKE '4%' THEN CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)) ELSE 0 END AS '桶數'
            //,CASE WHEN TA006 LIKE '4%' THEN CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1)) ELSE 0 END AS '箱數'

            FASTSQL.AppendFormat(@"    
                                SELECT TA003 AS '製令日期' ,TA001 AS '製令別',TA002 AS '製令編號',TA021 AS '生產線別',TA006 AS '品號',TA034 AS '品名',TA035 AS '規格',TA015 AS '預計產量',TA017 AS '實際產出',TA007 AS '單位',TA029 AS '備註',MB023,MB198
                                ,CASE WHEN MB198='2' THEN  CONVERT(NVARCHAR,DATEADD(DAY,-1,DATEADD(MONTH,MB023,TA003)),112) ELSE CONVERT(NVARCHAR,DATEADD(DAY,-1,DATEADD(DAY,MB023,TA003)),112) END AS '有效日期'
                                ,[ERPINVMB].[PCT] AS '比例'
                                ,[ERPINVMB].[ALLERGEN]  AS '過敏原'
                                ,[ERPINVMB].[SPEC] AS '餅體'
                                ,CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)) AS '桶數'
                                ,CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1)) AS '箱數'
                                ,MOCTA.UDF01 AS '順序'
                                ,ISNULL(MC004,1) MC004
                                ,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010
                                FROM [TK].dbo.MOCTA
                                LEFT JOIN [TK].dbo.INVMB ON MB001=TA006
                                LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].MB001=TA006
                                LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                                LEFT JOIN [TK].dbo.BOMMD ON MD035 LIKE '%箱%' AND MD003 LIKE '2%' AND MD007>1 AND MD001=TA006
                                WHERE TA003='{0}' 
                                ORDER BY TA003,TA021,TA001,TA002    
                                ", SDAY);

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2(string SDAY)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\製令明細表2020V2.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL2(SDAY);
            Table.SelectCommand = SQL;

            report1.Preview = previewControl1;
            report1.Show();
        }


        public string SETFASETSQL2(string SDAY)
        {
            StringBuilder FASTSQL = new StringBuilder();

            //,CASE WHEN TA006 NOT LIKE '4%' THEN CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)) ELSE 0 END AS '桶數'
            //,CASE WHEN TA006 LIKE '4%' THEN CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1)) ELSE 0 END AS '箱數'

            FASTSQL.AppendFormat(@"    
                                SELECT 
                                 [ID]
                                ,[REPORTMOCMANULINE].[MANULINE] AS '生產線別'
                                ,[REPORTMOCMANULINE].[LOTNO] AS 'LOTNO'
                                ,[REPORTMOCMANULINE].[TA001] AS '製令別'
                                ,[REPORTMOCMANULINE].[TA002] AS '製令編號'
                                ,CONVERT(NVARCHAR,[REPORTMOCMANULINE].[TA003],112) AS '製令日期'
                                ,[REPORTMOCMANULINE].[TA006] AS '品號'
                                ,[REPORTMOCMANULINE].[TA007] AS '單位'
                                ,[REPORTMOCMANULINE].[TA015] AS '預計產量'
                                ,[REPORTMOCMANULINE].[TA017] AS '實際產出'
                                ,[REPORTMOCMANULINE].[MB002] AS '品名'
                                ,[REPORTMOCMANULINE].[MB003] AS '規格'
                                ,[REPORTMOCMANULINE].[PCTS] AS '比例'
                                ,[REPORTMOCMANULINE].[SEQ] AS '順序'
                                ,[REPORTMOCMANULINE].[ALLERGEN]  AS '過敏原'
                                ,[REPORTMOCMANULINE].[COOKIES] AS '餅體'
                                ,[REPORTMOCMANULINE].[BARS] AS '桶數'
                                ,[REPORTMOCMANULINE].[BOXS] AS '箱數'
                                ,CONVERT(NVARCHAR,[REPORTMOCMANULINE].[VDATES],112) AS '有效日期'
                                ,[REPORTMOCMANULINE].[COMMENT] AS '備註'
                                ,MOCTA.TA026 AS '訂單別'
                                ,MOCTA.TA027 AS '訂單號'
                                ,TC053  AS '客戶'
                                FROM [TKMOC].[dbo].[REPORTMOCMANULINE]
                                LEFT JOIN [TK].dbo.MOCTA ON [REPORTMOCMANULINE].TA001=MOCTA.[TA001] AND [REPORTMOCMANULINE].[TA002]=MOCTA.[TA002]
                                LEFT JOIN [TK].dbo.COPTC ON TC001= TA026 AND TC002=TA027 
                                WHERE CONVERT(NVARCHAR,[REPORTMOCMANULINE].TA003,112)='{0}'   
                                ORDER BY [REPORTMOCMANULINE].TA003,[MANULINE],[REPORTMOCMANULINE].TA001,[REPORTMOCMANULINE].TA002   
                                ", SDAY);

            return FASTSQL.ToString();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            SETCODE();

            //textBox3.Text = SEARCHMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"));

            //if (!string.IsNullOrEmpty(textBox3.Text))
            //{
            //    SETCODE();
            //}
        }

        public string SEARCHMOCLOTNO(string MOCDATES)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT  [MOCDATES],[LOTNO]
                                    FROM [TKMOC].[dbo].[MOCLOTNO]
                                    WHERE [MOCDATES]='{0}'
                                    ", MOCDATES);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["TEMPds1"].Rows[0]["LOTNO"].ToString().Trim();
                }
                else
                {
                    return null;

                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDDELETEMOCLOTNO(string MOCDATES,string LOTNO)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(@" 
                                    DELETE [TKMOC].[dbo].[MOCLOTNO] WHERE [MOCDATES]='{0}'

                                    INSERT INTO  [TKMOC].[dbo].[MOCLOTNO] ( [MOCDATES],[LOTNO])
                                    VALUES ('{0}','{1}')
                                    ", MOCDATES, LOTNO);



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


        public void SearchMOCTA(string TA003)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            StringBuilder Query = new StringBuilder();

      
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" 
                                    SELECT TA021 AS '生產線別',TA001  AS '製令單',TA002  AS '製令單號',UDF01  AS '生產順序'
                                    FROM [TK].dbo.MOCTA
                                    WHERE TA003='{0}'
                                    ORDER BY TA021,TA001,TA002,UDF01
                                    ", TA003);

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

                        dataGridView1.Columns["生產線別"].ReadOnly = true;
                        dataGridView1.Columns["製令單"].ReadOnly = true;
                        dataGridView1.Columns["製令單號"].ReadOnly = true;
                       

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

        public void UPDATEMOCTA()
        {
            sbSql.Clear();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string TA001 = row.Cells[1].Value.ToString().Trim();
                string TA002 = row.Cells[2].Value.ToString().Trim();
                string UDF01 = row.Cells[3].Value.ToString().Trim();

                if(!string.IsNullOrEmpty(UDF01))
                {
                    sbSql.AppendFormat(@" UPDATE  [TK].dbo.MOCTA SET UDF01='{2}' WHERE TA001='{0}' AND  TA002='{1}'", TA001, TA002,UDF01);
                    sbSql.AppendFormat(@" ");
                }
                else
                {
                    sbSql.AppendFormat(@" UPDATE  [TK].dbo.MOCTA SET UDF01=NULL WHERE TA001='{0}' AND  TA002='{1}'", TA001, TA002, UDF01);
                    sbSql.AppendFormat(@" ");
                }
            }


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();


                sbSql.AppendFormat(@" ");


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



        public void ADDREPORTMOCMANULINE(string LOTNO ,string TA003)
        {
            sbSql.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();


                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKMOC].[dbo].[REPORTMOCMANULINE]
                                    ([ID],[MANULINE],[LOTNO],[TA001],[TA002],[TA003],[TA006],[TA007],[TA015],[TA017],[MB002],[MB003],[PCTS],[SEQ],[ALLERGEN],[COOKIES],[BARS],[BOXS],[VDATES],[COMMENT])

                                    SELECT NEWID(),TA021,'{0}',TA001,TA002,TA003,TA006,TA007,TA015,TA017,TA034,TA035,[ERPINVMB].[PCT],MOCTA.UDF01,[ERPINVMB].[ALLERGEN] ,[ERPINVMB].[SPEC] ,CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)),CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1)),
                                    CASE WHEN MB198='2' THEN  CONVERT(NVARCHAR,DATEADD(DAY,-1,DATEADD(MONTH,MB023,TA003)),112) ELSE CONVERT(NVARCHAR,DATEADD(DAY,-1,DATEADD(DAY,MB023,TA003)),112) END
                                    ,TA029
                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=TA006
                                    LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].MB001=TA006
                                    LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                                    LEFT JOIN [TK].dbo.BOMMD ON MD035 LIKE '%箱%' AND MD003 LIKE '2%' AND MD007>1 AND MD001=TA006
                                    WHERE [TA001]+[TA002] NOT IN (SELECT [TA001]+[TA002] FROM [TKMOC].[dbo].[REPORTMOCMANULINE] WHERE TA003='{1}')
                                    AND TA003='{1}' 
                                    ORDER BY TA003,TA021,TA001,TA002     
                                    ", LOTNO, TA003);


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

        public void SEARCHMOCLOTNO2(string MOCDATES)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            StringBuilder Query = new StringBuilder();


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" 
                                    SELECT 
                                    [MOCDATES] AS '日期'
                                    ,[LOTNO] AS '代碼'
                                    FROM [TKMOC].[dbo].[MOCLOTNO]
                                    WHERE [MOCDATES]='{0}'
                                    ", MOCDATES);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds1.Tables["ds1"];
                        dataGridView3.AutoResizeColumns();

                        dataGridView3.Columns["日期"].ReadOnly = true;
                        dataGridView3.Columns["代碼"].ReadOnly = true;
                     

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


        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count >= 1)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox9.Text = row.Cells["日期"].Value.ToString();
                  

                }
                else
                {                   
                    textBox9.Text = null;

                }
            }
        }

        public void DELETEMOCLOTNO(string MOCDATES)
        {
            sbSql.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();


                sbSql.AppendFormat(@" 
                                   DELETE [TKMOC].[dbo].[MOCLOTNO]
                                   WHERE [MOCDATES]='{0}'
                                    ", MOCDATES);


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

        public void SEARCHREPORTMOCMANULINE(string TA003)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            StringBuilder Query = new StringBuilder();


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" 
                                    SELECT 
                                     [ID]
                                    ,[MANULINE] AS '生產線別'
                                    ,[LOTNO] AS 'LOTNO'
                                    ,[TA001] AS '製令別'
                                    ,[TA002] AS '製令編號'
                                    ,CONVERT(NVARCHAR,[TA003],112) AS '製令日期'
                                    ,[TA006] AS '品號'
                                    ,[TA007] AS '單位'
                                    ,[TA015] AS '預計產量'
                                    ,[TA017] AS '實際產出'
                                    ,[MB002] AS '品名'
                                    ,[MB003] AS '規格'
                                    ,[PCTS] AS '比例'
                                    ,[SEQ] AS '順序'
                                    ,[ALLERGEN]  AS '過敏原'
                                    ,[COOKIES] AS '餅體'
                                    ,[BARS] AS '桶數'
                                    ,[BOXS] AS '箱數'
                                    ,CONVERT(NVARCHAR,[VDATES],112) AS '有效日期'
                                    ,[COMMENT] AS '備註'
                                    FROM [TKMOC].[dbo].[REPORTMOCMANULINE]
                                    WHERE CONVERT(NVARCHAR,TA003,112)='{0}' 
                                    ORDER BY TA003,[MANULINE],TA001,TA002   
                                    ", TA003);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds1.Tables["ds1"];
                        dataGridView4.AutoResizeColumns();



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

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.Rows.Count >= 1)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    textBox10.Text = row.Cells["ID"].Value.ToString();


                }
                else
                {
                    textBox10.Text = null;

                }
            }
        }

        public void DELREPORTMOCMANULINE(string ID)
        {
            sbSql.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();


                sbSql.AppendFormat(@" 
                                  DELETE [TKMOC].[dbo].[REPORTMOCMANULINE]
                                  WHERE [ID]='{0}'
                                    ", ID);


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

        public void DELREPORTMOCMANULINE2(string TA003)
        {
            sbSql.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();


                sbSql.AppendFormat(@" 
                                  DELETE [TKMOC].[dbo].[REPORTMOCMANULINE]
                                  WHERE CONVERT(NVARCHAR,TA003,112)='{0}'
                                    ", TA003);


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

        public void SETCODE()
        {
            DataSet ds = new DataSet();
            string yyyyMMdd = dateTimePicker1.Value.ToString("yyyyMMdd");
            string MM = Convert.ToUInt32(yyyyMMdd.Substring(4,2)).ToString();//除0開頭
            string d1 = yyyyMMdd.Substring(6,1);
            string d2 = yyyyMMdd.Substring(7,1);
            string CODE = "";
            string CODE1= "";
            string CODE2 = "";
            string CODE3 = "";

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT [ID],[CODE] FROM [TKMOC].[dbo].[MOCCODE]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);


                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    foreach (DataRow od in ds.Tables["ds"].Rows)
                    {
                        if(MM.Equals(od["ID"].ToString()))
                        {
                            CODE1 = od["CODE"].ToString();
                        }
                        if (d1.Equals(od["ID"].ToString()))
                        {
                            CODE2 = od["CODE"].ToString();
                        }
                        if (d2.Equals(od["ID"].ToString()))
                        {
                            CODE3 = od["CODE"].ToString();
                        }
                    }

                    textBox3.Text = CODE1 + CODE2 + CODE3;
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SETNULL1()
        {
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";

            textBox12.ReadOnly = false;
            //textBox13.ReadOnly = false;
            textBox14.ReadOnly = false;
        }

        public void SETNULL2()
        {
            textBox12.ReadOnly = true;
            //textBox13.ReadOnly = true;
            textBox14.ReadOnly = true;
        }
        public void SETNULL3()
        {
            textBox12.ReadOnly = false;
            //textBox13.ReadOnly = false;
            textBox14.ReadOnly = false;
        }

        public void SEARCHMOCHALFPRODUCTDBOXS()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                if (string.IsNullOrEmpty(textBox11.Text))
                {
                    sbSql.AppendFormat(@"  
                                SELECT [MOCHALFPRODUCTDBOXS].[MB001] AS '品號',[MB002] AS '品名',[NUMS] AS '箱重',[BOXS] AS '箱數'
                                FROM [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS],[TK].dbo.[INVMB]
                                WHERE [MOCHALFPRODUCTDBOXS].[MB001]=[INVMB].[MB001]
                                 ");
                }
                else if (!string.IsNullOrEmpty(textBox11.Text))
                {
                    sbSql.AppendFormat(@"  
                                SELECT [MOCHALFPRODUCTDBOXS].[MB001] AS '品號',[MB002] AS '品名',[NUMS] AS '箱重',[BOXS] AS '箱數'
                                FROM [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS],[TK].dbo.[INVMB]
                                WHERE [MOCHALFPRODUCTDBOXS].[MB001]=[INVMB].[MB001]
                                AND ([MOCHALFPRODUCTDBOXS].[MB001] LIKE '%{0}%' OR [MB002] LIKE '%{0}%')
                                 ", textBox11.Text);
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView5.DataSource = ds1.Tables["ds1"];
                    dataGridView5.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
              
                else
                {
                    dataGridView5.DataSource = null;
                }

            }
            catch
            {

            }
            finally
            {

            }

        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    textBox12.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox14.Text = row.Cells["箱重"].Value.ToString().Trim();


                }
                else
                {


                }
            }
        }


        public void ADDMOCHALFPRODUCTDBOXS(string MB001, string NUMS, string BOXS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    INSERT INTO[TKMOC].[dbo].[MOCHALFPRODUCTDBOXS]
                                    ([MB001],[NUMS],[BOXS])
                                    VALUES('{0}',{1},{2})
                                        ", MB001, NUMS, 1);


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

        public void UPDATEMOCHALFPRODUCTDBOXS(string MB001, string NUMS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                   UPDATE [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS]
                                    SET [NUMS]={1}
                                    WHERE [MB001]='{0}'
                                        ", MB001, NUMS);


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

        public void DELETEMOCHALFPRODUCTDBOXS(string MB001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                   DELETE [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS]
                                    WHERE [MB001]='{0}'
                                        ", MB001);


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

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            textBox13.Text = SERCHINVMB(textBox12.Text.Trim());
        }
        public string SERCHINVMB(string MB001)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT MB002 FROM [TK].dbo.INVMB WHERE MB001='{0}'
                                    ", MB001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["MB002"].ToString().Trim();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {

            }
        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ADDERPINVMB();
            Search();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            UPDATEERPINVMB();
            Search();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker1.Value.ToString("yyyyMMdd"));

            //if(!string.IsNullOrEmpty(textBox3.Text.Trim()))
            //{
            //    ADDDELETEMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"), textBox3.Text.Trim());
            //}

            //textBox3.Text = SEARCHMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"));

            //SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"),textBox3.Text.Trim());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SearchMOCTA(dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        private void button5_Click(object sender, EventArgs e)
        {
            UPDATEMOCTA();
            SearchMOCTA(dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox3.Text.Trim()))
            {
                ADDDELETEMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"), textBox3.Text.Trim());
                textBox3.Text = SEARCHMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"));

                ADDREPORTMOCMANULINE(textBox3.Text.Trim(), dateTimePicker1.Value.ToString("yyyyMMdd"));

                MessageBox.Show("完成");
            }

            
        }

      
        private void button9_Click(object sender, EventArgs e)
        {
            SEARCHMOCLOTNO2(dateTimePicker4.Value.ToString("yyyyMMdd"));
        }
        private void button10_Click(object sender, EventArgs e)
        {
            DELETEMOCLOTNO(textBox9.Text);
            SEARCHMOCLOTNO2(dateTimePicker4.Value.ToString("yyyyMMdd"));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SEARCHREPORTMOCMANULINE(dateTimePicker3.Value.ToString("yyyyMMdd"));
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //DELREPORTMOCMANULINE(textBox10.Text);
            DELREPORTMOCMANULINE2(dateTimePicker3.Value.ToString("yyyyMMdd"));
            SEARCHREPORTMOCMANULINE(dateTimePicker3.Value.ToString("yyyyMMdd"));
        }

        private void button12_Click(object sender, EventArgs e)
        {
            SEARCHMOCHALFPRODUCTDBOXS();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SETNULL1();
            STATUS = "ADD";
        }

        private void button14_Click(object sender, EventArgs e)
        {
            SETNULL3();
            STATUS = "UPDATE";
        }

        private void button16_Click(object sender, EventArgs e)
        {

            SETNULL2();

            if (STATUS.Equals("ADD"))
            {
                ADDMOCHALFPRODUCTDBOXS(textBox12.Text.Trim(), textBox14.Text.Trim(), "1");
            }
            else if (STATUS.Equals("UPDATE"))
            {
                UPDATEMOCHALFPRODUCTDBOXS(textBox12.Text.Trim(), textBox14.Text.Trim());
            }

            STATUS = "";
            SEARCHMOCHALFPRODUCTDBOXS();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除嗎?", "要刪除嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETEMOCHALFPRODUCTDBOXS(textBox12.Text.Trim());

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }


            STATUS = "";
            SEARCHMOCHALFPRODUCTDBOXS();
        }




        #endregion

      
    }
}
