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
using TKITDLL;

namespace TKMOC
{
    public partial class frmREPORTGENBAKING : Form
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
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        int rownum = 0;
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        int ROWS = 0;
        int TA017 = 0;
        decimal CHECKROWS = 0;
        string MB001 = null;
        decimal BOXNUM = 0;
        public Report report1 { get; private set; }

        public frmREPORTGENBAKING()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCHMOCTA(DateTime dt, DateTime dt2)
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
                                    SELECT TA001 AS '製令',TA002 AS '製令號',TA006 AS '品號',TA034 AS '品名' ,TA015 AS '預計產量',TA007 AS '單位',TA035  AS '規格',TA029 '備註'
                                    FROM [TK].dbo.MOCTA 
                                    WHERE TA003>='{0}' AND TA003<='{1}' 
                                    AND TA001 IN ('A513') 
                                    ORDER BY TA001,TA002,TA034 
                                    ", dt.ToString("yyyyMMdd"), dt2.ToString("yyyyMMdd"));

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
            MB001 = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["製令"].Value.ToString();
                    textBox2.Text = row.Cells["製令號"].Value.ToString();
                    textBox3.Text = row.Cells["備註"].Value.ToString();
                    MB001 = row.Cells["品號"].Value.ToString();

                    SEARCHREPORTGENDETAIL(textBox1.Text, textBox2.Text);
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    MB001 = null;

                }
            }
        }

        public void ADD_REPORTGENBAKING(string TA001, string TA002, string TA017)
        {
            ROWS = SERACHERPINVMB(TA001, TA002, TA017);
            int fianl = 1;
            int BOSNUMS = 0;

            if (ROWS > 0)
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

                    sbSql.AppendFormat("DELETE  [TKMOC].[dbo].[REPORTGENBAKING]");

                    string BOARDNUM = ds2.Tables["ds2"].Rows[0]["BOARDNUM"].ToString();
                    int BOARDNUMGENNUM = Convert.ToInt32(Convert.ToDecimal(BOARDNUM));

                    if (CHECKROWS == ROWS)
                    {
                        for (int i = 1; i <= ROWS; i++)
                        {
                            sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[REPORTGENBAKING]");
                            sbSql.AppendFormat(" ([TA001],[TA002],[YEARS],[MONTHS],[DAYS],[MB001],[MB002],[MB003],[GENNUM],[BORADNUM])");
                            sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',{9})", ds2.Tables["ds2"].Rows[0]["TA001"].ToString(), ds2.Tables["ds2"].Rows[0]["TA002"].ToString(), ds2.Tables["ds2"].Rows[0]["YEARS"].ToString(), ds2.Tables["ds2"].Rows[0]["MONTHS"].ToString(), ds2.Tables["ds2"].Rows[0]["DAYS"].ToString(), ds2.Tables["ds2"].Rows[0]["MB001"].ToString(), ds2.Tables["ds2"].Rows[0]["MB002"].ToString(), ds2.Tables["ds2"].Rows[0]["MB003"].ToString(), BOARDNUMGENNUM.ToString() + 'A', i);
                            sbSql.AppendFormat(" ");
                        }
                    }
                    else
                    {
                        while (CHECKROWS > 1)
                        {
                            BOSNUMS = BOSNUMS + 1;

                            sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[REPORTGENBAKING]");
                            sbSql.AppendFormat(" ([TA001],[TA002],[YEARS],[MONTHS],[DAYS],[MB001],[MB002],[MB003],[GENNUM],[BORADNUM])");
                            sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',{9})", ds2.Tables["ds2"].Rows[0]["TA001"].ToString(), ds2.Tables["ds2"].Rows[0]["TA002"].ToString(), ds2.Tables["ds2"].Rows[0]["YEARS"].ToString(), ds2.Tables["ds2"].Rows[0]["MONTHS"].ToString(), ds2.Tables["ds2"].Rows[0]["DAYS"].ToString(), ds2.Tables["ds2"].Rows[0]["MB001"].ToString(), ds2.Tables["ds2"].Rows[0]["MB002"].ToString(), ds2.Tables["ds2"].Rows[0]["MB003"].ToString(), BOARDNUMGENNUM.ToString() + 'A', BOSNUMS);
                            sbSql.AppendFormat(" ");
                            fianl = fianl + 1;

                            CHECKROWS = CHECKROWS - 1;
                        }

                        decimal GENNUMS = 0;

                        if (MB001.Substring(0, 1).Equals("4"))
                        {
                            decimal INPUT = Convert.ToDecimal(textBox4.Text);
                            decimal OTHERS = Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOXNUM"].ToString()) * Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOARDNUM"].ToString()) * BOSNUMS;
                            GENNUMS = Math.Ceiling((INPUT - OTHERS) / Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOXNUM"].ToString()));


                        }
                        else if (MB001.Substring(0, 1).Equals("3"))
                        {
                            decimal INPUT = Convert.ToDecimal(textBox4.Text);
                            decimal OTHERS = Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOXNUM"].ToString()) * Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOARDNUM"].ToString()) * BOSNUMS;
                            GENNUMS = Math.Ceiling((INPUT - OTHERS) / Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOXNUM"].ToString()));

                            //GENNUMS = Math.Ceiling((Convert.ToDecimal(Convert.ToDecimal(textBox4.Text) - (Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOXNUM"].ToString()) * Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOARDNUM"].ToString()) * (BOSNUMS)))) / (Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOXNUM"].ToString())));
                        }



                        sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[REPORTGENBAKING]");
                        sbSql.AppendFormat(" ([TA001],[TA002],[YEARS],[MONTHS],[DAYS],[MB001],[MB002],[MB003],[GENNUM],[BORADNUM])");
                        sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8} A',{9})", ds2.Tables["ds2"].Rows[0]["TA001"].ToString(), ds2.Tables["ds2"].Rows[0]["TA002"].ToString(), ds2.Tables["ds2"].Rows[0]["YEARS"].ToString(), ds2.Tables["ds2"].Rows[0]["MONTHS"].ToString(), ds2.Tables["ds2"].Rows[0]["DAYS"].ToString(), ds2.Tables["ds2"].Rows[0]["MB001"].ToString(), ds2.Tables["ds2"].Rows[0]["MB002"].ToString(), ds2.Tables["ds2"].Rows[0]["MB003"].ToString(), Convert.ToInt32(GENNUMS), fianl);
                        sbSql.AppendFormat(" ");
                    }


                    sbSql.AppendFormat("  ");
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

        }
        public int SERACHERPINVMB(string TA001, string TA002, string TA017)
        {
            BOXNUM = 0;

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

                sbSql.AppendFormat(@"  SELECT [TA001],[TA002],SUBSTRING(TA002,1,4) AS 'YEARS',SUBSTRING(TA002,5,2) AS 'MONTHS',SUBSTRING(TA002,7,2) AS 'DAYS',INVMB.[MB001],INVMB.[MB002],INVMB.[MB003],{0} AS 'TA017',CONVERT(DECIMAL(16,4),{0})/(CONVERT(DECIMAL(16,4),[ERPINVMB].BOXNUM)*CONVERT(DECIMAL(16,4),[ERPINVMB].BOARDNUM)) AS 'ROWS',[PROCESSTIME],[BOXNUM],[ERPINVMB].BOARDNUM  ", TA017);
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MOCTA.TA006");
                sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].MB001=INVMB.MB001 ");
                sbSql.AppendFormat(@"  WHERE TA001='{0}' AND TA002='{1}'", TA001, TA002);
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
                        CHECKROWS = Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["ROWS"].ToString());
                        BOXNUM = Convert.ToDecimal(ds2.Tables["ds2"].Rows[0]["BOXNUM"].ToString());
                        return Convert.ToInt32(CHECKROWS);
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

        public void ADD_REPORTGENDETAILBAKING(string TA001, string TA002, string COMMENTS)
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
                                    INSERT INTO [TKMOC].[dbo].[REPORTGENDETAILBAKING]
                                    ([TA001],[TA002],[COMMENTS])
                                    VALUES ('{0}','{1}','{2}')
                                    ", TA001, TA002, COMMENTS);

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

        public void SETFASTREPORT2(string TA001, string TA002, string COMMENT, string NUM)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\生產入庫單(自動)(烘焙).frx");

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


            FASTSQL.AppendFormat(@"    
                                SELECT [TA001]  AS '製令',[TA002] AS '製令號',[YEARS] AS '年',[MONTHS] AS '月',[DAYS] AS '日',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[GENNUM]  AS '已生產量' ,[BORADNUM]  AS '版數'     
                                FROM [TKMOC].[dbo].[REPORTGENBAKING]    
                                WHERE TA001='{0}' AND TA002='{1}'
                                ORDER BY [TA001],[TA002],[BORADNUM]  
                                ", TA001, TA002);

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT(string TA001, string TA002, string COMMENT, string NUM)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\生產入庫單(烘焙).frx");

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



            FASTSQL.AppendFormat(@"     
                                SELECT TA001 AS '製令',TA002 AS '製令號',SUBSTRING(TA002,1,4) AS '年',SUBSTRING(TA002,5,2) AS '月',SUBSTRING(TA002,7,2) AS '日',TA034 AS '品名',MB003 AS '規格'
                                ,MB003 AS '規格',TA017 AS '已生產量' 
                                FROM [TK].dbo.MOCTA
                                LEFT JOIN [TK].dbo.INVMB ON MB001=TA006
                                WHERE TA001='{0}' AND TA002='{1}'
                                ", TA001, TA002);

            return FASTSQL.ToString();
        }

        public void SEARCHREPORTGENDETAIL(string TA001, string TA002)
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
                                    SELECT TOP 1  COMMENTS 
                                    FROM [TKMOC].[dbo].[REPORTGENDETAILBAKING]
                                    WHERE TA001='{0}' AND TA002='{1}'
                                    ORDER BY CDATES DESC
                                    ", TA001, TA002);

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    textBox9.Text = null;
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        textBox9.Text = ds4.Tables["ds4"].Rows[0]["COMMENTS"].ToString();
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCTA(dateTimePicker1.Value, dateTimePicker2.Value);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            int NUM = 0;
            int N;
            ADD_REPORTGENBAKING(textBox1.Text, textBox2.Text, textBox4.Text);
            ADD_REPORTGENDETAILBAKING(textBox1.Text, textBox2.Text, textBox4.Text);

            if (ROWS > 0)
            {
                SETFASTREPORT2(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
            }
            else
            {
                if (BOXNUM > 0 && (int.TryParse(textBox4.Text, out N)))
                {
                    decimal CALNUM = Convert.ToDecimal(textBox4.Text) / BOXNUM;
                    NUM = Convert.ToInt32(Math.Round(CALNUM, 0, MidpointRounding.AwayFromZero));

                    SETFASTREPORT(textBox1.Text, textBox2.Text, textBox3.Text, NUM.ToString() + " A");
                }
                else
                {
                    SETFASTREPORT(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
                }


            }

            SEARCHREPORTGENDETAIL(textBox1.Text, textBox2.Text);
        }

        #endregion

        
    }
}
