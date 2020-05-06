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
    public partial class frmREPORTGEN : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;

        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        int ROWS=0;
        int TA017=0;

        public Report report1 { get; private set; }

        public frmREPORTGEN()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SEARCHMOCTA(DateTime dt,DateTime dt2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT TA001 AS '製令',TA002 AS '製令號',TA034 AS '品名' ,TA015 AS '預計產量',TA007 AS '單位',TA035  AS '規格',TA029 '備註'");
                sbSql.AppendFormat(@" FROM [TK].dbo.MOCTA ");
                sbSql.AppendFormat(@" WHERE TA003>='{0}' AND TA003<='{1}' ", dt.ToString("yyyyMMdd"), dt2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@" AND TA001 IN ('A510','A511','A512','A521','A522') ");
                sbSql.AppendFormat(@" ORDER BY TA001,TA002,TA034 ");
                sbSql.AppendFormat(@"   ");
                sbSql.AppendFormat(@"  ");

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
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["製令"].Value.ToString();
                    textBox2.Text = row.Cells["製令號"].Value.ToString();
                    textBox3.Text = row.Cells["備註"].Value.ToString();
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;

                }
            }
        }

        public void SETFASTREPORT(string TA001,string TA002, string COMMENT, string NUM)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\生產入庫單.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL(TA001, TA002);
            Table.SelectCommand = SQL;

            report1.SetParameterValue("P1", COMMENT);
            report1.SetParameterValue("P2", NUM);

            report1.Preview = previewControl1;
            report1.Show();
        }


        public string SETFASETSQL(string TA001, string TA002)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

           
            FASTSQL.AppendFormat(@" SELECT TA001 AS '製令',TA002 AS '製令號',SUBSTRING(TA002,1,4) AS '年',SUBSTRING(TA002,5,2) AS '月',SUBSTRING(TA002,7,2) AS '日',TA034 AS '品名',MB003 AS '規格'");
            FASTSQL.AppendFormat(@"  ,MB003 AS '規格',TA017 AS '已生產量' ");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.MOCTA");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON MB001=TA006");
            FASTSQL.AppendFormat(@"  WHERE TA001='{0}' AND TA002='{1}'",TA001,TA002);
            FASTSQL.AppendFormat(@"     ");

            return FASTSQL.ToString();
        }

        public int SERACHERPINVMB(string TA001, string TA002,string TA017)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [TA001],[TA002],SUBSTRING(TA002,1,4) AS 'YEARS',SUBSTRING(TA002,5,2) AS 'MONTHS',SUBSTRING(TA002,7,2) AS 'DAYS',INVMB.[MB001],INVMB.[MB002],INVMB.[MB003],{0} AS 'TA017',CONVERT(INT,ROUND({0}/[ERPINVMB].BOARDNUM,0)) AS 'ROWS',[ERPINVMB].BOARDNUM  ", Convert.ToInt32(TA017));
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MOCTA.TA006");
                sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].MB001=INVMB.MB001 ");
                sbSql.AppendFormat(@"  WHERE TA001='{0}' AND TA002='{1}'",TA001,TA002);
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    return 0;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        return Convert.ToInt32(ds2.Tables["ds2"].Rows[0]["ROWS"].ToString());
                    }

                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void ADDREPORTGEN(string TA001,string TA002,string TA017)
        {
            ROWS = SERACHERPINVMB(TA001, TA002, TA017);
            if(ROWS > 0)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat("DELETE  [TKMOC].[dbo].[REPORTGEN]");
                    for (int i=1;i<= ROWS;i++)
                    {
                        sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[REPORTGEN]");
                        sbSql.AppendFormat(" ([TA001],[TA002],[YEARS],[MONTHS],[DAYS],[MB001],[MB002],[MB003],[GENNUM],[BORADNUM])");
                        sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}',{8},{9})", ds2.Tables["ds2"].Rows[0]["TA001"].ToString(), ds2.Tables["ds2"].Rows[0]["TA002"].ToString(), ds2.Tables["ds2"].Rows[0]["YEARS"].ToString(), ds2.Tables["ds2"].Rows[0]["MONTHS"].ToString(), ds2.Tables["ds2"].Rows[0]["DAYS"].ToString(), ds2.Tables["ds2"].Rows[0]["MB001"].ToString(), ds2.Tables["ds2"].Rows[0]["MB002"].ToString(), ds2.Tables["ds2"].Rows[0]["MB003"].ToString(), ds2.Tables["ds2"].Rows[0]["TA017"].ToString(), i);
                        sbSql.AppendFormat(" ");
                    }

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
            
        }

        public void SETFASTREPORT2(string TA001, string TA002, string COMMENT, string NUM)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\生產入庫單(自動).frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL2(TA001, TA002);
            Table.SelectCommand = SQL;

            report1.SetParameterValue("P1", COMMENT);
            report1.SetParameterValue("P2", NUM);

            report1.Preview = previewControl1;
            report1.Show();
        }


        public string SETFASETSQL2(string TA001, string TA002)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();


            FASTSQL.AppendFormat(@" SELECT [TA001]  AS '製令',[TA002] AS '製令號',[YEARS] AS '年',[MONTHS] AS '月',[DAYS] AS '日',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[GENNUM]  AS '已生產量' ,[BORADNUM]  AS '版數'     ");
            FASTSQL.AppendFormat(@" FROM [TKMOC].[dbo].[REPORTGEN]    ");
            FASTSQL.AppendFormat(@" WHERE TA001='{0}' AND TA002='{1}'", TA001, TA002);
            FASTSQL.AppendFormat(@" ORDER BY [TA001],[TA002],[BORADNUM]   ");
            FASTSQL.AppendFormat(@"     ");

            return FASTSQL.ToString();
        }

        public void Search()
        {
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

                sbSql.AppendFormat(@" SELECT [MB001] AS '品號',[MB002]  AS '品名',[MB003]  AS '規格',[BOXNUM]  AS '箱數',[BOARDNUM]  AS '板數'");
                sbSql.AppendFormat(@" FROM [TKMOC].[dbo].[ERPINVMB] ");
                sbSql.AppendFormat(@" WHERE 1=1 ");
                sbSql.AppendFormat(@" {0}", Query.ToString());
                sbSql.AppendFormat(@"  ORDER BY [MB001]");
                sbSql.AppendFormat(@" ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {

                        
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds3.Tables["ds3"];
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
                sbSql.AppendFormat(@"INSERT INTO [TKMOC].dbo.ERPINVMB (MB001,MB002,MB003,[PROCESSNUM] ,[PROCESSTIME],[BOXNUM],[BOARDNUM]) ");
                sbSql.AppendFormat(@" SELECT MB001,MB002,MB003,0,0,0,0 FROM [TK].dbo.INVMB WITH (NOLOCK) WHERE (MB001 LIKE '4%' OR MB001 LIKE '3%' ) AND MB001 NOT IN (SELECT MB001 FROM [TKMOC].dbo.ERPINVMB WITH (NOLOCK) )");
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
                    numericUpDown1.Value =Convert.ToInt32(row.Cells["箱數"].Value.ToString());
                    numericUpDown1.Value = Convert.ToInt32(row.Cells["板數"].Value.ToString());

                }
                else
                {
                    textBox6.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                    numericUpDown1.Value = 0;
                    numericUpDown1.Value = 0;

                }
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCTA(dateTimePicker1.Value, dateTimePicker2.Value);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDREPORTGEN(textBox1.Text, textBox2.Text,textBox4.Text);

            if(ROWS>0)
            {
                SETFASTREPORT2(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
            }
            else
            {
                SETFASTREPORT(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
            }
            
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ADDERPINVMB();
        }


        #endregion

       
    }
}
