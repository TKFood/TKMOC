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
using FastReport;
using FastReport.Data;
using System.Collections;

namespace TKMOC
{
    public partial class frmDEVINVMB : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();


        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        int result;

       
        string ID;
       

        public frmDEVINVMB()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH()
        {
            
            StringBuilder ISLOSED = new StringBuilder();
            StringBuilder SLQURY = new StringBuilder();

            if(comboBox1.Text.Equals("否"))
            {
                ISLOSED.AppendFormat(@" AND [ISCLOSED] IN ('N') ");
            }
            else if (comboBox1.Text.Equals("是"))
            {
                ISLOSED.AppendFormat(@" AND [ISCLOSED] IN ('Y') ");
            }
            else
            {
                ISLOSED.AppendFormat(@" AND [ISCLOSED] IN ('Y','N') ");
            }

            if(!string.IsNullOrEmpty(textBox1.Text))
            {
                SLQURY.AppendFormat(@" AND OLDMB001 LIKE '%{0}%'", textBox1.Text);
            }
            else
            {
                SLQURY.AppendFormat(@" ");
            }

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT [NO] AS '校稿編號',CONVERT(NVARCHAR,[SDATES],111) AS '起始日期',[OLDMB001] AS '原品號',[OLDMB002] AS '物料名稱',[NEWMB001] AS '新品號',[NEWMB002] AS '新物料名稱',CONVERT(NVARCHAR,[PURDATES],111) AS '預計發包日',[ISUSED] AS '用完改版',[ISSCRAPPED] AS '報廢',[ISCLOSED] AS '是否結案'");
                sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA001=OLDMB001 ) AS '原品號庫存'");
                sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA001=NEWMB001 ) AS '新品號庫存'");
                sbSql.AppendFormat(@"  ,[ID] ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[DEVINVMB]");
                sbSql.AppendFormat(@"  WHERE 1=1");
                sbSql.AppendFormat(@"  {0}", ISLOSED.ToString());
                sbSql.AppendFormat(@"  {0}", SLQURY.ToString());
                sbSql.AppendFormat(@"  ORDER BY OLDMB001");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];

                        dataGridView1.AutoResizeColumns();
                        dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;


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

        public void SEARCH2()
        {

            StringBuilder ISLOSED = new StringBuilder();
            StringBuilder SLQURY = new StringBuilder();

            if (comboBox2.Text.Equals("否"))
            {
                ISLOSED.AppendFormat(@" AND [ISCLOSED] IN ('N') ");
            }
            else if (comboBox2.Text.Equals("是"))
            {
                ISLOSED.AppendFormat(@" AND [ISCLOSED] IN ('Y') ");
            }
            else
            {
                ISLOSED.AppendFormat(@" AND [ISCLOSED] IN ('Y','N') ");
            }

            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                SLQURY.AppendFormat(@" AND OLDMB001 LIKE '%{0}%'", textBox2.Text);
            }
            else
            {
                SLQURY.AppendFormat(@" ");
            }

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT [NO] AS '校稿編號',CONVERT(NVARCHAR,[SDATES],111) AS '起始日期',[OLDMB001] AS '原品號',[OLDMB002] AS '物料名稱',[NEWMB001] AS '新品號',[NEWMB002] AS '新物料名稱',CONVERT(NVARCHAR,[PURDATES],111) AS '預計發包日',[ISUSED] AS '用完改版',[ISSCRAPPED] AS '報廢',[ISCLOSED] AS '是否結案'");
                sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA001=OLDMB001 ) AS '原品號庫存'");
                sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA001=NEWMB001 ) AS '新品號庫存'");
                sbSql.AppendFormat(@"  ,[ID] ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[DEVINVMB]");
                sbSql.AppendFormat(@"  WHERE 1=1");
                sbSql.AppendFormat(@"  {0}", ISLOSED.ToString());
                sbSql.AppendFormat(@"  {0}", SLQURY.ToString());
                sbSql.AppendFormat(@"  ORDER BY OLDMB001");
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
                        dataGridView2.DataSource = ds2.Tables["ds2"];

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;


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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SEARCH2();
        }

        #endregion


    }
}
