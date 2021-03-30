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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString()))
            {
                textBox1.Text = comboBox1.SelectedValue.ToString();
            }
            else
            {
                textBox1.Text = "";
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString().Trim()))
            {
                SEARCH1(comboBox1.SelectedValue.ToString().Trim());

                decimal CAL = Convert.ToDecimal(textBox2.Text);
                SEARCH2(comboBox1.SelectedValue.ToString().Trim(), CAL);
            }
                
        }
        #endregion

      
    }
}
