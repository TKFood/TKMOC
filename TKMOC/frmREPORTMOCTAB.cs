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
using System.Xml;

namespace TKMOC
{
    public partial class frmREPORTMOCTAB : Form
    {
        //測試ID = "";
        //正式ID =""
        //測試DB DBNAME = "UOFTEST";
        //正式DB DBNAME = "UOF";
        /// <summary>
        /// 生產排程確認表
        /// </summary>
        string ID1 = "6ecb6b2b-72db-4431-9782-93045391a562";
        /// <summary>
        /// 生產排程確認表說明
        /// </summary>
        string ID2 = "63bde26a-54ec-465d-b525-6f5fad42fae7";
        string DBNAME = "UOF";


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

            if (!string.IsNullOrEmpty(textBox3.Text))
            {
                SEARCHUOF(textBox3.Text);
                SEARCHUOF2(textBox3.Text);
            }

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

                sbSql.AppendFormat(@" SELECT [MB001] AS '品號',[MB002]  AS '品名',[MB003]  AS '規格',[PCT]  AS '比例',[ALLERGEN]  AS '過敏原',[SPEC] AS '餅體',[ORI] AS '素別' ");
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
                
                sbSql.AppendFormat(@"  
                                    UPDATE [TKMOC].[dbo].[ERPINVMB]
                                    SET [PCT]='{1}',[ALLERGEN]='{2}',[SPEC]='{3}' ,[ORI]='{4}'
                                    WHERE [MB001]='{0}'
                                    ", textBox6.Text.ToString().Trim(), textBox1.Text.ToString().Trim(), textBox2.Text.ToString().Trim(), textBox4.Text.ToString().Trim(), textBox15.Text.ToString().Trim());

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
                                ,[REPORTMOCMANULINE].[ORI] AS '素別'
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
            SEARCHVDATES(dateTimePicker1.Value.ToString("yyyyMMdd"));
            if(!string.IsNullOrEmpty(textBox3.Text))
            {
                SEARCHUOF(textBox3.Text);
                SEARCHUOF2(textBox3.Text);
            }
            

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

        public void ADDDELETEMOCLOTNO(string MOCDATES, string LOTNO, string VDATES)
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

                                    INSERT INTO  [TKMOC].[dbo].[MOCLOTNO] ( [MOCDATES],[LOTNO],[VDATES])
                                    VALUES ('{0}','{1}','{2}')
                                    ", MOCDATES, LOTNO, VDATES);



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

                if (!string.IsNullOrEmpty(UDF01))
                {
                    sbSql.AppendFormat(@" UPDATE  [TK].dbo.MOCTA SET UDF01='{2}' WHERE TA001='{0}' AND  TA002='{1}'", TA001, TA002, UDF01);
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



        public void ADDREPORTMOCMANULINE(string LOTNO, string TA003)
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
                                    ([ID],[MANULINE],[LOTNO],[TA001],[TA002],[TA003],[TA006],[TA007],[TA015],[TA017],[MB002],[MB003],[PCTS],[SEQ],[ALLERGEN],[COOKIES],[BARS],[BOXS],[VDATES],[COMMENT],[ORI])

                                    SELECT NEWID(),TA021,'{0}',TA001,TA002,TA003,TA006,TA007,TA015,TA017,TA034,TA035,[ERPINVMB].[PCT],MOCTA.UDF01,[ERPINVMB].[ALLERGEN] ,[ERPINVMB].[SPEC] ,CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)),CASE WHEN ISNULL(1/[MOCHALFPRODUCTDBOXS].NUMS*[MOCHALFPRODUCTDBOXS].BOXS,0)>0  THEN CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1))*(1/[MOCHALFPRODUCTDBOXS].NUMS*[MOCHALFPRODUCTDBOXS].BOXS) ELSE CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1)) END,
                                    CASE WHEN MB198='2' THEN  CONVERT(NVARCHAR,DATEADD(DAY,-1,DATEADD(MONTH,MB023,TA003)),112) ELSE CONVERT(NVARCHAR,DATEADD(DAY,-1,DATEADD(DAY,MB023,TA003)),112) END
                                    ,TA029
                                    ,[ERPINVMB].[ORI]
                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=TA006
                                    LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].MB001=TA006
                                    LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                                    LEFT JOIN [TK].dbo.BOMMD ON MD035 LIKE '%箱%' AND MD003 LIKE '2%' AND MD007>1 AND MD001=TA006 AND MD003 NOT  IN ('201001237')
                                    LEFT JOIN [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS] ON [MOCHALFPRODUCTDBOXS].MB001=TA006
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
            string MM = Convert.ToUInt32(yyyyMMdd.Substring(4, 2)).ToString();//除0開頭
            string d1 = yyyyMMdd.Substring(6, 1);
            string d2 = yyyyMMdd.Substring(7, 1);
            string CODE = "";
            string CODE1 = "";
            string CODE2 = "";
            string CODE3 = "";
            string VDATES = "";

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

                        if (MM.Equals(od["ID"].ToString()))
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

        public void SEARCHVDATES(string MOCDATES)
        {
            DataSet ds = new DataSet();
            string VDATES = null;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT [MOCDATES] ,[LOTNO],[VDATES] FROM [TKMOC].[dbo].[MOCLOTNO]
                                    WHERE [MOCDATES]='{0}'
                                    ", MOCDATES);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);


                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    VDATES = ds.Tables["ds"].Rows[0]["VDATES"].ToString();

                    if (!string.IsNullOrEmpty(VDATES))
                    {
                        dateTimePicker5.Value = Convert.ToDateTime(VDATES.Substring(0, 4) + '/' + VDATES.Substring(4, 2) + '/' + VDATES.Substring(6, 2));
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
                                    INSERT INTO [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS]
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

        public void ADDTB_WKF_EXTERNAL_TASK(string sday)
        {

            //人員指定 190006
            string ACCOUNT = "190006";
            string CODE = textBox3.Text.Trim();     
            string VDATES = dateTimePicker5.Value.ToString("yyyyMMdd");

            DataTable DTUPFDEP = SEARCHUOFDEP(ACCOUNT);
            DataTable DT = SEARCHDB(sday);

            string account = DTUPFDEP.Rows[0]["ACCOUNT"].ToString();
            string groupId = DTUPFDEP.Rows[0]["GROUP_ID"].ToString();
            string jobTitleId = DTUPFDEP.Rows[0]["TITLE_ID"].ToString();
            string fillerName = DTUPFDEP.Rows[0]["NAME"].ToString();
            string fillerUserGuid = DTUPFDEP.Rows[0]["USER_GUID"].ToString();

            string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
            string DEPNO = DTUPFDEP.Rows[0]["DEPNO"].ToString();

            string EXTERNAL_FORM_NBR = CODE;

            int rowscounts = 0;

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");

            //正式的id
            Form.SetAttribute("formVersionId", ID1);

            Form.SetAttribute("urgentLevel", "2");
            //加入節點底下
            xmlDoc.AppendChild(Form);

            ////建立節點Applicant
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");
            Applicant.SetAttribute("account", account);
            Applicant.SetAttribute("groupId", groupId);
            Applicant.SetAttribute("jobTitleId", jobTitleId);
            //加入節點底下
            Form.AppendChild(Applicant);

            //建立節點 Comment
            XmlElement Comment = xmlDoc.CreateElement("Comment");
            Comment.InnerText = "申請者意見";
            //加入至節點底下
            Applicant.AppendChild(Comment);

            //建立節點 FormFieldValue
            XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
            //加入至節點底下
            Form.AppendChild(FormFieldValue);

            //建立節點FieldItem
            //ID 表單編號	
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "ID");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA003	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA003");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["製令日期"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //PCODE	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "PCODE");
            FieldItem.SetAttribute("fieldValue", CODE);
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //VDATES	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "VDATES");
            FieldItem.SetAttribute("fieldValue", VDATES);
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //DETAILS	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "DETAILS");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點 DataGrid
            XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
            //DataGrid 加入至 DETAILS 節點底下
            XmlNode DETAILS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='DETAILS']");
            DETAILS.AppendChild(DataGrid);

            foreach (DataRow od in DT.Rows)
            {
                // 新增 Row
                XmlElement Row = xmlDoc.CreateElement("Row");
                Row.SetAttribute("order", (rowscounts).ToString());

                //Row	MANULINE
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "MANULINE");
                Cell.SetAttribute("fieldValue", od["生產線別"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TA001
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TA001");
                Cell.SetAttribute("fieldValue", od["製令別"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TA002
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TA002");
                Cell.SetAttribute("fieldValue", od["製令編號"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TA015
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TA015");
                Cell.SetAttribute("fieldValue", od["預計產量"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TA007
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TA007");
                Cell.SetAttribute("fieldValue", od["單位"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	BARS
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "BARS");
                Cell.SetAttribute("fieldValue", od["桶數"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	BOXS
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "BOXS");
                Cell.SetAttribute("fieldValue", od["箱數"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	MB002
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "MB002");
                Cell.SetAttribute("fieldValue", od["品名"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	MB003
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "MB003");
                Cell.SetAttribute("fieldValue", od["規格"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	COOKIES
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "COOKIES");
                Cell.SetAttribute("fieldValue", od["餅體"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	PCTS
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "PCTS");
                Cell.SetAttribute("fieldValue", od["比例"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	SEQ
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "SEQ");
                Cell.SetAttribute("fieldValue", od["順序"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	ALLERGEN
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "ALLERGEN");
                Cell.SetAttribute("fieldValue", od["過敏原"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	ORI
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "ORI");
                Cell.SetAttribute("fieldValue", od["素別"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
              

                //Row	COMMENT
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "COMMENT");
                Cell.SetAttribute("fieldValue", od["備註"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TC053
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TC053");
                Cell.SetAttribute("fieldValue", od["客戶"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	SIGNS
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "SIGNS");
                Cell.SetAttribute("fieldValue", "");
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                rowscounts = rowscounts + 1;

                XmlNode DataGridS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='DETAILS']/DataGrid");
                DataGridS.AppendChild(Row);
            }


                ////用ADDTACK，直接啟動起單
                //ADDTACK(Form);

                //ADD TO DB
                string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@" INSERT INTO [{0}].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS,EXTERNAL_FORM_NBR)
                                        VALUES (NEWID(),@XML,2,'{1}')
                                        ", DBNAME, EXTERNAL_FORM_NBR);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                    MessageBox.Show("新增成功");

                }
            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ADDTB_WKF_EXTERNAL_TASK2(string sday)
        {

            //人員指定 190006
            string ACCOUNT = "190006";
            string CODE = textBox3.Text.Trim();
            string VDATES = dateTimePicker5.Value.ToString("yyyyMMdd");
            string REASONS= textBox18.Text.Trim();

            DataTable DTUPFDEP = SEARCHUOFDEP(ACCOUNT);
            DataTable DT = SEARCHDB(sday);

            string account = DTUPFDEP.Rows[0]["ACCOUNT"].ToString();
            string groupId = DTUPFDEP.Rows[0]["GROUP_ID"].ToString();
            string jobTitleId = DTUPFDEP.Rows[0]["TITLE_ID"].ToString();
            string fillerName = DTUPFDEP.Rows[0]["NAME"].ToString();
            string fillerUserGuid = DTUPFDEP.Rows[0]["USER_GUID"].ToString();

            string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
            string DEPNO = DTUPFDEP.Rows[0]["DEPNO"].ToString();

            string EXTERNAL_FORM_NBR = "REASONS-"+CODE;

            int rowscounts = 0;

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");

            //正式的id
            Form.SetAttribute("formVersionId", ID2);

            Form.SetAttribute("urgentLevel", "2");
            //加入節點底下
            xmlDoc.AppendChild(Form);

            ////建立節點Applicant
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");
            Applicant.SetAttribute("account", account);
            Applicant.SetAttribute("groupId", groupId);
            Applicant.SetAttribute("jobTitleId", jobTitleId);
            //加入節點底下
            Form.AppendChild(Applicant);

            //建立節點 Comment
            XmlElement Comment = xmlDoc.CreateElement("Comment");
            Comment.InnerText = "申請者意見";
            //加入至節點底下
            Applicant.AppendChild(Comment);

            //建立節點 FormFieldValue
            XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
            //加入至節點底下
            Form.AppendChild(FormFieldValue);

            //建立節點FieldItem
            //ID 表單編號	
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "ID");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA003	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA003");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["製令日期"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //PCODE	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "PCODE");
            FieldItem.SetAttribute("fieldValue", CODE);
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //VDATES	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "VDATES");
            FieldItem.SetAttribute("fieldValue", VDATES);
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //VDATES	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "REASONS");
            FieldItem.SetAttribute("fieldValue", REASONS);
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //DETAILS	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "DETAILS");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點 DataGrid
            XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
            //DataGrid 加入至 DETAILS 節點底下
            XmlNode DETAILS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='DETAILS']");
            DETAILS.AppendChild(DataGrid);

            foreach (DataRow od in DT.Rows)
            {
                // 新增 Row
                XmlElement Row = xmlDoc.CreateElement("Row");
                Row.SetAttribute("order", (rowscounts).ToString());

                //Row	MANULINE
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "MANULINE");
                Cell.SetAttribute("fieldValue", od["生產線別"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TA001
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TA001");
                Cell.SetAttribute("fieldValue", od["製令別"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TA002
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TA002");
                Cell.SetAttribute("fieldValue", od["製令編號"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TA015
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TA015");
                Cell.SetAttribute("fieldValue", od["預計產量"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TA007
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TA007");
                Cell.SetAttribute("fieldValue", od["單位"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	BARS
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "BARS");
                Cell.SetAttribute("fieldValue", od["桶數"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	BOXS
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "BOXS");
                Cell.SetAttribute("fieldValue", od["箱數"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	MB002
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "MB002");
                Cell.SetAttribute("fieldValue", od["品名"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	MB003
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "MB003");
                Cell.SetAttribute("fieldValue", od["規格"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	COOKIES
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "COOKIES");
                Cell.SetAttribute("fieldValue", od["餅體"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	PCTS
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "PCTS");
                Cell.SetAttribute("fieldValue", od["比例"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	SEQ
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "SEQ");
                Cell.SetAttribute("fieldValue", od["順序"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	ALLERGEN
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "ALLERGEN");
                Cell.SetAttribute("fieldValue", od["過敏原"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	ORI
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "ORI");
                Cell.SetAttribute("fieldValue", od["素別"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);


                //Row	COMMENT
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "COMMENT");
                Cell.SetAttribute("fieldValue", od["備註"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TC053
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TC053");
                Cell.SetAttribute("fieldValue", od["客戶"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	SIGNS
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "SIGNS");
                Cell.SetAttribute("fieldValue", "");
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                rowscounts = rowscounts + 1;

                XmlNode DataGridS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='DETAILS']/DataGrid");
                DataGridS.AppendChild(Row);
            }


            ////用ADDTACK，直接啟動起單
            //ADDTACK(Form);

            //ADD TO DB
            string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@" INSERT INTO [{0}].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS,EXTERNAL_FORM_NBR)
                                        VALUES (NEWID(),@XML,2,'{1}')
                                        ", DBNAME, EXTERNAL_FORM_NBR);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                    MessageBox.Show("新增成功");

                }
            }
            catch
            {

            }
            finally
            {

            }
        }

        public DataTable SEARCHUOFDEP(string ACCOUNT)
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
                                    SELECT 
                                    [GROUP_NAME] AS 'DEPNAME'
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID]+','+[GROUP_NAME]+',False' AS 'DEPNO'
                                    ,[TB_EB_USER].[USER_GUID] AS 'USER_GUID'
                                    ,[ACCOUNT] AS 'ACCOUNT'
                                    ,[NAME] AS 'NAME'
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID] AS'GROUP_ID'
                                    ,[TITLE_ID]   AS 'TITLE_ID'
                                    ,[GROUP_NAME] AS 'GROUP_NAME'
                                    ,[GROUP_CODE] AS'GROUP_CODE'
                                    FROM [192.168.1.223].[{0}].[dbo].[TB_EB_USER],[192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP],[192.168.1.223].[{0}].[dbo].[TB_EB_GROUP]
                                    WHERE [TB_EB_USER].[USER_GUID]=[TB_EB_EMPL_DEP].[USER_GUID]
                                    AND [TB_EB_EMPL_DEP].[GROUP_ID]=[TB_EB_GROUP].[GROUP_ID]
                                    AND ISNULL([TB_EB_GROUP].[GROUP_CODE],'')<>''
                                    AND [ACCOUNT]='{1}'
                              
                                    ", DBNAME, ACCOUNT);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

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

        public DataTable SEARCHDB(string SDAY)
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
                                ,[REPORTMOCMANULINE].[ORI] AS '素別'
                                FROM [TKMOC].[dbo].[REPORTMOCMANULINE]
                                LEFT JOIN [TK].dbo.MOCTA ON [REPORTMOCMANULINE].TA001=MOCTA.[TA001] AND [REPORTMOCMANULINE].[TA002]=MOCTA.[TA002]
                                LEFT JOIN [TK].dbo.COPTC ON TC001= TA026 AND TC002=TA027 
                                WHERE CONVERT(NVARCHAR,[REPORTMOCMANULINE].TA003,112)='{0}'   
                                AND [REPORTMOCMANULINE].[TA001] IN ('A510','A512')  
                                ORDER BY [REPORTMOCMANULINE].TA003,[MANULINE],[REPORTMOCMANULINE].TA001,[REPORTMOCMANULINE].TA002   

                                ", SDAY);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

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

        public void SEARCHUOF(string CODE)
        {
            textBox16.Text = null;
            textBox17.Text = null;

            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TOP 1 [TB_WKF_EXTERNAL_TASK].TASK_ID, EXTERNAL_FORM_NBR,[TB_WKF_EXTERNAL_TASK].DOC_NBR,CONVERT(NVARCHAR,MODIFY_TIME,112) MODIFY_TIMES,MODIFY_TIME
                                    FROM [UOF].[dbo].[TB_WKF_EXTERNAL_TASK],[UOF].[dbo].[TB_WKF_TASK] 
                                    WHERE [TB_WKF_EXTERNAL_TASK].TASK_ID=[TB_WKF_TASK].TASK_ID 
                                    AND ([TB_WKF_TASK].TASK_RESULT IN ('-1','0','3') OR ISNULL([TB_WKF_TASK].TASK_RESULT,'')='')
                                    AND EXTERNAL_FORM_NBR='{0}'
                                    ORDER BY MODIFY_TIME DESC

                                    ", CODE);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {

                    textBox16.Text = ds1.Tables["ds1"].Rows[0]["MODIFY_TIMES"].ToString();
                    textBox17.Text = ds1.Tables["ds1"].Rows[0]["DOC_NBR"].ToString();
                }
                else
                {
                    textBox16.Text = null;
                    textBox17.Text = null;
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

        public void SEARCHUOF2(string CODE)
        {
            textBox19.Text = null;
            
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                CODE = "REASONS-" + CODE;
                sbSql.AppendFormat(@"  
                                    SELECT TOP 1 [TB_WKF_EXTERNAL_TASK].TASK_ID, EXTERNAL_FORM_NBR,[TB_WKF_EXTERNAL_TASK].DOC_NBR,CONVERT(NVARCHAR,MODIFY_TIME,112) MODIFY_TIMES,MODIFY_TIME
                                    FROM [UOF].[dbo].[TB_WKF_EXTERNAL_TASK],[UOF].[dbo].[TB_WKF_TASK] 
                                    WHERE [TB_WKF_EXTERNAL_TASK].TASK_ID=[TB_WKF_TASK].TASK_ID 
                                    AND ([TB_WKF_TASK].TASK_RESULT IN ('-1','0','3') OR ISNULL([TB_WKF_TASK].TASK_RESULT,'')='')
                                    AND EXTERNAL_FORM_NBR='{0}'
                                    ORDER BY MODIFY_TIME DESC

                                    ", CODE);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {                                        
                    textBox19.Text = ds1.Tables["ds1"].Rows[0]["DOC_NBR"].ToString();
                }
                else
                {
                    textBox19.Text = null;
                   
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
            if (!string.IsNullOrEmpty(textBox3.Text))
            {
                SEARCHUOF(textBox3.Text);
                SEARCHUOF2(textBox3.Text);
            }

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
            if (!string.IsNullOrEmpty(textBox3.Text))
            {
                SEARCHUOF(textBox3.Text);
                SEARCHUOF2(textBox3.Text);
            }

            if (!string.IsNullOrEmpty(textBox3.Text.Trim()))
            {
                ADDDELETEMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"), textBox3.Text.Trim(), dateTimePicker5.Value.ToString("yyyyMMdd"));
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


        private void button17_Click(object sender, EventArgs e)
        {
            ADDTB_WKF_EXTERNAL_TASK(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox18.Text))
            {
                ADDTB_WKF_EXTERNAL_TASK2(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            else
            {
                MessageBox.Show("說明必填");
            }
           
        }

        #endregion


    }
}
