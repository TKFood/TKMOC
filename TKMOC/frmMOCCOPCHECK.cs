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
using System.Text.RegularExpressions;
using System.Globalization;

namespace TKMOC
{
    public partial class frmMOCCOPCHECK : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();

        int result;
        string tablename = null;
        string ID;

        public frmMOCCOPCHECK()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCHCOPTA()
        {
            StringBuilder SLQURY = new StringBuilder();
            StringBuilder SLQURY2 = new StringBuilder();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                SLQURY.Clear();

                if (checkBox2.Checked == true)
                {
                    SLQURY.AppendFormat(@"  AND TD001+TD002+TD003 NOT IN (SELECT [TA026]+[TA027]+[TA028] FROM [TK].dbo.MOCTA)");
                }
                if (checkBox3.Checked == true)
                {
                    SLQURY.AppendFormat(@"  AND TD001+TD002+TD003  NOT IN (SELECT [COPTA001]+[COPTA002]+[COPTA003] FROM [TKMOC].[dbo].[MOCCOPCHECK])");
                }

                sbSql.AppendFormat(@"  SELECT TC053 AS '客戶',TD013 AS '預交日',MV002 AS '業務',TD001 AS '訂單別',TD002 AS '訂單號',TD003 AS '訂單序號',TD004 AS '品號',TD005 AS '品名',(TD008-TD009+TD024-TD025) AS '需求量',TD010 AS '單位',TD008 AS '訂單量',TD009 AS '已交量',TD024 AS '贈品量',TD025 AS '已交贈品'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.CMSMV");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND MV001=TC006");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND TD001 NOT IN ('A228')");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  {0}", SLQURY.ToString());
                sbSql.AppendFormat(@"  {0}", SLQURY2.ToString());
       

                sbSql.AppendFormat(@"  ORDER BY TC053,TD013");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds"];

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

        #endregion

        #region BUTTON
        private void button11_Click(object sender, EventArgs e)
        {
            SEARCHCOPTA();
        }

        #endregion

    }
}
