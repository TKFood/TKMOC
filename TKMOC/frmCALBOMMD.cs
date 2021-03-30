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
            if(!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString().Trim()))
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString().Trim()))
            {
                SEARCH1(comboBox1.SelectedValue.ToString().Trim());

                decimal CAL = Convert.ToDecimal(textBox2.Text);
                SEARCH2(comboBox1.SelectedValue.ToString().Trim(), CAL);
                SEARCH3(comboBox1.SelectedValue.ToString().Trim(), CAL);
                SEARCH4(comboBox1.SelectedValue.ToString().Trim());
            }
                
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCH5(comboBox2.SelectedValue.ToString().Trim());
        }


        #endregion

       
    }
}
