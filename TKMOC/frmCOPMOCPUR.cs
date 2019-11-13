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

namespace TKMOC
{
    public partial class frmCOPMOCPUR : Form
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

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        string MID;
        string DID;

        public frmCOPMOCPUR()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCHCOP(DateTime dt1,DateTime dt2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

               
                sbSql.AppendFormat(@"  SELECT TD013 AS '預交日',TD001 AS '訂單',TD002 AS '訂單號',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
                sbSql.AppendFormat(@"  ,CONVERT(DECIMAL(18,3),(CASE WHEN MD002=TD010  THEN (TD008-TD009)*MD004/MD003 ELSE (TD008-TD009) END )) AS '數量'");
                sbSql.AppendFormat(@"  ,TD010 AS '單位',TC015 AS '單頭備註',TD020 AS '單身備註'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=MB001");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD004=MB001");
                sbSql.AppendFormat(@"  AND (TD004 LIKE '410%')");
                sbSql.AppendFormat(@"  AND (TD008-TD009)>0");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'",dt1.ToString("yyyyMMdd"), dt2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY TD013,TD001,TD002,TD004");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
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


        public void SEARCHCOPMOCPUR(string MID,string DID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                
                sbSql.AppendFormat(@"  SELECT [MID] AS '來源單別',[DID] AS '來源單號',[TA001] AS '採購單',[TA002] AS '採購單號'");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[COPMOCPUR]");
                sbSql.AppendFormat(@"  WHERE [MID]='{0}' AND [DID]='{1}'",MID,DID);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");


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

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["訂單"].Value.ToString();
                    textBox2.Text = row.Cells["訂單號"].Value.ToString();

                    SEARCHCOPMOCPUR(textBox1.Text, textBox2.Text);
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;

                }
            }
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHCOP(dateTimePicker1.Value,dateTimePicker2.Value);
        }
        #endregion

       
    }
}
