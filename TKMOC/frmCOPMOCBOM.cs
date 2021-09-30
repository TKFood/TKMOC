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
using TKITDLL;

namespace TKMOC
{
    public partial class frmCOPMOCBOM : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();

        string tablename = null;
        int rownum = 0;

        public frmCOPMOCBOM()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    //dataGridView1.Columns.Clear();
                    ds.Clear();

                    adapter.Fill(ds, "ds");
                    sqlConn.Close();

                    if (ds.Tables["ds"].Rows.Count == 0)
                    {
                        dataGridView1.DataSource = null;
                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                    }
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

            }

        }
        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();

            STR.AppendFormat(@"  SELECT TD013 AS '預交日',TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',(TD008-TD009)  AS '數量',TD010 AS '訂單單位'");
            STR.AppendFormat(@"  ,(CASE WHEN TD010=MD002 THEN ((TD008-TD009)*MD004/MD003) ELSE (TD008-TD009) END) AS '生產數量'");
            STR.AppendFormat(@"  ,MB004 AS '單位'");
            STR.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.INVMB,[TK].dbo.COPTD");
            STR.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004");
            STR.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
            STR.AppendFormat(@"  AND MB001=TD004");
            STR.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            STR.AppendFormat(@"  AND TD021='Y' ");
            STR.AppendFormat(@"  AND TD016='N'");
            STR.AppendFormat(@"  ORDER BY TD004,TD013");
            STR.AppendFormat(@"  ");


            return STR;
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        #endregion
    }

}
