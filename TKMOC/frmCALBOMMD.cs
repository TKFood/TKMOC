using System;
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

namespace TKMOC
{
    public partial class frmCALBOMMD : Form
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

        //水麵倍數
        decimal CAL1;
        //油酥倍數
        decimal CAL2;
        //油酥所需的水面倍數
        decimal CAL3;
        //水麵顆數
        decimal CALNUM1;
        //油酥顆數
        decimal CALNUM2;
        //油酥所需的水面顆數
        decimal CALNUM3;

        public frmCALBOMMD()
        {
            InitializeComponent();

            comboBox1load();
        }



        #region FUNCTION
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD003,MB002 FROM [TKMOC].[dbo].[MOCSEPECIALCAL]  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD003", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD003";
            comboBox1.DisplayMember = "MB002";
            sqlConn.Close();

            textBox1.Text = "";

        }

        public void comboBox2load(string MD003)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT MD001,MB002
                                FROM [TK].dbo.BOMMD
                                LEFT JOIN [TK].dbo.INVMB ON MB001=BOMMD.MD001
                                WHERE MD003='{0}'
                                AND MB002 NOT LIKE '%暫停%'
                                ORDER BY MD001
                                ", MD003);

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD001";
            comboBox2.DisplayMember = "MB002";
            sqlConn.Close();

          

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SETNULL();

            if (!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString().Trim()))
            {
                textBox1.Text = comboBox1.SelectedValue.ToString().Trim();

                comboBox2load(comboBox1.SelectedValue.ToString().Trim());
            }
            else
            {
                textBox1.Text = "";
            }

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(comboBox2.SelectedValue.ToString().Trim()))
            {
                textBox5.Text = comboBox2.SelectedValue.ToString().Trim();

              
            }
            else
            {
                textBox5.Text = "";
            }
        }

        //一桶水面-先算出中筋一桶的倍率=66
        public void SEARCH1(string MD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();
            
                sbSql.AppendFormat(@"  
                                    SELECT [MOCSEPECIALCAL].[MD003],66/BOMMD.MD006 AS 'CAL'
                                    FROM [TKMOC].[dbo].[MOCSEPECIALCAL],[TK].dbo.BOMMD
                                    WHERE [MOCSEPECIALCAL].MD003=BOMMD.MD001
                                    AND BOMMD.MD003 LIKE '1%'
                                    AND [MOCSEPECIALCAL].[MD003]='{0}'
                                    AND BOMMD.MD003='101001001'
                                    ORDER BY [MOCSEPECIALCAL].[MD003]
                                    ", MD003);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    textBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["CAL"].ToString();
                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    //dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {
                    
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

        //一桶水面-用「先算出中筋一桶的倍率=66」算其他料的用量
        public void SEARCH2(string MD003,decimal CAL)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                   SELECT BOMMD.MD003 AS '元件品號',MB002  AS '品名',CONVERT(decimal(16,4),BOMMD.MD006*({1})) AS '用量' ,BOMMD.MD007  AS '底數',BOMMD.MD008  AS '損耗率%',BOMMD.MD001  AS '主件品號'
                                    FROM[TK].dbo.BOMMD
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=BOMMD.MD003
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('101001009')
                                    AND BOMMD.MD001='{0}'
                                    ORDER BY BOMMD.MD003
                                    ", MD003, CAL);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {                   
                    //dataGridView1.Rows.Clear();
                    dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        //--一桶水面-合計用量
        public void SEARCH3(string MD003, decimal CAL)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT BOMMD.MD001,SUM(BOMMD.MD006*({1})) AS 'SUMCALMD006' 
                                    ,(SELECT TOP 1 [MOCSEPECIALCAL].[WATERNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003 =BOMMD.MD001 ) AS 'WATERNUM'
                                    ,(SUM(BOMMD.MD006*({1}))/((SELECT TOP 1 [MOCSEPECIALCAL].[WATERNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003=BOMMD.MD001 ) )) AS 'WATERNUMS'
                                    FROM[TK].dbo.BOMMD
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('101001009')
                                    AND BOMMD.MD001='{0}'
                                    GROUP BY BOMMD.MD001
                                    ", MD003, CAL);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    textBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["WATERNUMS"].ToString();

                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    //dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        public void SEARCH4(string MD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT [MD003],[MB002],[WATERNUMS],[OILNUMS]
                                    FROM [TKMOC].[dbo].[MOCSEPECIALCAL]
                                    WHERE [MD003]='{0}'
                                    ", MD003);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    textBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["WATERNUMS"].ToString();

                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    //dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        //--一桶油酥-先算出低筋一桶的倍率=66
        public void SEARCH5(string MD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT BOMMD.[MD001],66/BOMMD.MD006 AS 'CAL'
                                    FROM [TK].dbo.BOMMD
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003='101001002'
                                    AND BOMMD.MD001='{0}'

                                    ORDER BY BOMMD.[MD001]
                                    ", MD003);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    textBox6.Text = ds1.Tables["TEMPds1"].Rows[0]["CAL"].ToString();

                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    //dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        //--一桶油酥-用「先算出低筋一桶的倍率=66」算其他料的用量
        public void SEARCH6(string MD003,decimal CAL)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT BOMMD.MD003 AS '元件品號',MB002  AS '品名',CONVERT(decimal(16,4),BOMMD.MD006*({1}) ) AS '用量' ,BOMMD .MD007  AS '底數',BOMMD.MD008  AS '損耗率%',BOMMD.MD001 AS '主件品號'
                                    FROM[TK].dbo.BOMMD
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=BOMMD.MD003
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('101001009')
                                    AND BOMMD.MD001='{0}'
                                    ORDER BY BOMMD.MD003
                                    ", MD003, CAL);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView2.DataSource = ds1.Tables["TEMPds1"];
                    dataGridView2.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        //--一桶水面-合計用量
        public void SEARCH7(string MD003,decimal CAL)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT BOMMD.MD001,SUM(BOMMD.MD006*({1})) AS 'SUMCALMD006' 
                                    ,(SELECT TOP 1 [MOCSEPECIALCAL].[OILNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003 IN (SELECT MD003 FROM [TK].dbo.BOMMD MD WHERE MD.MD001=BOMMD.MD001)) AS 'OILNUM'
                                    ,(SUM(BOMMD.MD006*{1})/((SELECT TOP 1 [MOCSEPECIALCAL].[OILNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003 IN (SELECT MD003 FROM [TK].dbo.BOMMD MD WHERE MD.MD001=BOMMD.MD001)) )) AS 'OILNUMS'
                                    FROM[TK].dbo.BOMMD
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('101001009')
                                    AND BOMMD.MD001='{0}'
                                    GROUP BY BOMMD.MD001
                                    ", MD003, CAL);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    textBox7.Text = ds1.Tables["TEMPds1"].Rows[0]["OILNUMS"].ToString();

                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    //dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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


        public void SEARCH8(string MD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT [MD003],[MB002],[WATERNUMS],[OILNUMS]
                                    FROM [TKMOC].[dbo].[MOCSEPECIALCAL]
                                    WHERE [MD003]='{0}'
                                    ", MD003);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    textBox8.Text = ds1.Tables["TEMPds1"].Rows[0]["OILNUMS"].ToString();

                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    //dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        //--一桶油酥所需的水面原料
        public void SEARCH9(string MD003, decimal CAL1, decimal CAL2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT BOMMD.MD003 AS '元件品號',MB002  AS '品名',CONVERT(decimal(16,4),BOMMD.MD006*({1})*({2})) AS '用量' ,BOMMD .MD007  AS '底數',BOMMD.MD008  AS '損耗率%',BOMMD.MD001 AS '主件品號'
                                    FROM[TK].dbo.BOMMD
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=BOMMD.MD003
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('101001009')
                                    AND BOMMD.MD001='{0}'
                                    ORDER BY BOMMD.MD003
                                    ", MD003, CAL1, CAL2);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView3.DataSource = ds1.Tables["TEMPds1"];
                    dataGridView3.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        //--一桶油酥所需的水面原料SUM
        public void SEARCH10(string MD003, decimal CAL1, decimal CAL2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT BOMMD.MD001,SUM(BOMMD.MD006*({1})*({2})) AS 'SUMCALMD006' 
                                    ,(SELECT TOP 1 [MOCSEPECIALCAL].[WATERNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003 =BOMMD.MD001 ) AS 'WATERNUM'
                                    ,(SUM(BOMMD.MD006*({1})*({2}))/((SELECT TOP 1 [MOCSEPECIALCAL].[WATERNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003=BOMMD.MD001 ) )) AS 'WATERNUMS'

                                    FROM[TK].dbo.BOMMD
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('101001009')
                                    AND BOMMD.MD001='{0}'
                                    GROUP BY BOMMD.MD001
                                    ", MD003, CAL1, CAL2);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    textBox10.Text = ds1.Tables["TEMPds1"].Rows[0]["WATERNUMS"].ToString();

                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    //dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        //輸入桶數就可得知所需水面原料
        public void SEARCH11(string MD003, decimal CAL1, decimal CAL3, string MD001, decimal CAL2, decimal WORKNUMS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT 類型,元件品號,品名,用量
                                    FROM(
                                    SELECT '水面' AS '類型', BOMMD.MD003 AS '元件品號',MB002  AS '品名',CONVERT(decimal(16,4),BOMMD.MD006*({1})*({2})) AS '用量' ,BOMMD .MD007  AS '底數',BOMMD.MD008  AS '損耗率%',BOMMD.MD001 AS '主件品號'
                                    FROM[TK].dbo.BOMMD
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=BOMMD.MD003
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('101001009')
                                    AND BOMMD.MD001='{0}'
                                    UNION ALL
                                    
                                    SELECT '油酥' AS '類型',BOMMD.MD003 AS '元件品號',MB002  AS '品名',CONVERT(decimal(16,4),BOMMD.MD006*({4})*({5}) ) AS '用量' ,BOMMD .MD007  AS '底數',BOMMD.MD008  AS '損耗率%',BOMMD.MD001 AS '主件品號'
                                    FROM[TK].dbo.BOMMD
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=BOMMD.MD003
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('101001009')
                                    AND BOMMD.MD001='{3}'
                                   
                                    ) AS TEMP
                                    ORDER BY 類型,元件品號
                                    ", MD003, CAL1, CAL3, MD001, CAL2, WORKNUMS);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView4.DataSource = ds1.Tables["TEMPds1"];
                    dataGridView4.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        //--輸入桶數就可得知所需水面原料SUM
        public void SEARCH12(string MD003, decimal CAL1, decimal CAL2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT SUM(CONVERT(decimal(16,4),BOMMD.MD006*({1})*({2}))) AS '用量' 
                                    ,(SELECT TOP 1 [MOCSEPECIALCAL].[WATERNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003 =BOMMD.MD001 ) AS 'WATERNUM'
                                    ,(SUM(BOMMD.MD006*({1})*({2}))/((SELECT TOP 1 [MOCSEPECIALCAL].[WATERNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003=BOMMD.MD001 ) )) AS 'WATERNUMS'

                                    FROM[TK].dbo.BOMMD
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=BOMMD.MD003
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('{0}')
                                    AND BOMMD.MD001='3010000115'
                                    GROUP BY BOMMD.MD001
                                    ", MD003, CAL1, CAL2);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    textBox14.Text = ds1.Tables["TEMPds1"].Rows[0]["WATERNUMS"].ToString();

                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    //dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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

        public void SEARCHMOCSEPECIALCAL()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT [MD003] AS '品號',[MB002] AS '品名',[WATERNUMS] AS '水麵重',[OILNUMS] AS '油酥重'
                                    FROM [TKMOC].[dbo].[MOCSEPECIALCAL]
                                    ORDER BY [MD003]
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {

                    dataGridView5.Rows.Clear();
                    dataGridView5.DataSource = ds1.Tables["TEMPds1"];
                    dataGridView5.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else
                {

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
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;

            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            dataGridView4.DataSource = null;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString().Trim()))
            {
                SEARCH1(comboBox1.SelectedValue.ToString().Trim());

                CAL1 = Convert.ToDecimal(textBox2.Text);
                SEARCH2(comboBox1.SelectedValue.ToString().Trim(), CAL1);
                SEARCH3(comboBox1.SelectedValue.ToString().Trim(), CAL1);
                SEARCH4(comboBox1.SelectedValue.ToString().Trim());

                CALNUM1 = Convert.ToDecimal(textBox3.Text);
            }
                
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //SEARCH5(comboBox2.SelectedValue.ToString().Trim());
            //CAL2 = Convert.ToDecimal(textBox6.Text);

            textBox6.Text = "1";
            CAL2 = Convert.ToDecimal(textBox6.Text);

            SEARCH6(comboBox2.SelectedValue.ToString().Trim(), CAL2);
            SEARCH7(comboBox2.SelectedValue.ToString().Trim(), CAL2);
            SEARCH8(comboBox1.SelectedValue.ToString().Trim());

            CALNUM2 = Convert.ToDecimal(textBox7.Text);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(CALNUM1 > 0 && CALNUM2 > 0)
            {
                textBox9.Text = (CALNUM2 / CALNUM1).ToString();
                CAL3 = (CALNUM2 / CALNUM1);
                CALNUM3 = CALNUM1 * CAL3;
                SEARCH9(comboBox1.SelectedValue.ToString().Trim(),CAL1, CAL3);
                SEARCH10(comboBox1.SelectedValue.ToString().Trim(), CAL1, CAL3);

            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox11.Text))
            {
                decimal WORKNUMS = Convert.ToDecimal(textBox11.Text);

                if(WORKNUMS>0&& CALNUM3 > 0)
                {
                    textBox12.Text = (CALNUM3 * WORKNUMS).ToString();
                }
                if (WORKNUMS > 0 && CALNUM3 > 0 && CALNUM1 > 0)
                {
                    textBox13.Text = (CALNUM3 * WORKNUMS/ CALNUM1).ToString();
                    CAL3 = (CALNUM3 * WORKNUMS / CALNUM1);
                }

                SEARCH11(comboBox1.SelectedValue.ToString().Trim(), CAL1, CAL3, comboBox2.SelectedValue.ToString().Trim(), CAL2, WORKNUMS);

                SEARCH12(comboBox1.SelectedValue.ToString().Trim(), CAL1, CAL3);
            }
            
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SEARCHMOCSEPECIALCAL();
        }

        #endregion


    }
}
