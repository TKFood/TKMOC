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


        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        int rownum = 0;
        DataSet ds1 = new DataSet();

        public Report report1 { get; private set; }

        public frmREPORTMOCTAB()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void frmREPORTMOCTAB_Load(object sender, EventArgs e)
        {
            textBox3.Text = SEARCHMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }
        public void Search()
        {         
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            StringBuilder Query = new StringBuilder();

            if (!string.IsNullOrEmpty(textBox5.Text.ToString()))
            {
                Query.AppendFormat(@" AND MB001 LIKE '{0}%' ", textBox5.Text.ToString());
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
                                  ,[PCT] AS '比例'
                                ,[ALLERGEN]  AS '過敏原'
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox3.Text = SEARCHMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"));
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
            if(!string.IsNullOrEmpty(textBox3.Text.Trim()))
            {
                ADDDELETEMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"), textBox3.Text.Trim());
            }

            textBox3.Text = SEARCHMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"));

            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"),textBox3.Text.Trim());
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

        #endregion


    }
}
