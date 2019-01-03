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

namespace TKMOC
{
    public partial class frmRECOPSTILL : Form
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
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        public frmRECOPSTILL()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1();
            Report report1 = new Report();
            report1.Load(@"REPORT\訂單未出完且未結案報表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT TC053 AS '客戶',TD013 AS '預交日',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD008 AS '訂單量',TD009 AS '出貨量',TD024 AS '贈品量',TD025 AS '贈品已交量',(TD008-TD009+TD024-TD025) AS '總未出貨量',TD010 AS '單位',TD001 AS '訂單',TD002 AS '單號',TD003 AS '序號'");
            SB.AppendFormat(" FROM [TK].dbo.COPTD,[TK].dbo.COPTC");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TD004 LIKE '4%'");
            SB.AppendFormat(" AND (TD008-TD009+TD024-TD025)>0");
            SB.AppendFormat(" AND TD021='Y'");
            SB.AppendFormat(" AND TD016='N'");
            SB.AppendFormat(" AND TC001 IN ('A221', 'A222','A223','A228')");
            SB.AppendFormat(" ORDER BY TC001,TC053,TD013,TD004");
            SB.AppendFormat(" ");



            return SB;

        }
        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2();
            Report report2 = new Report();
            report2.Load(@"REPORT\訂單未出完且未結案的品號報表.frx");

            report2.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report2.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report2.Preview = previewControl2;
            report2.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',SUM(TD008) AS '訂單量',SUM(TD009)  AS '出貨量',SUM(TD024)  AS '贈品量',SUM(TD025)  AS '贈品已交量',SUM((TD008-TD009+TD024-TD025)) AS '總未出貨量',TD010 AS '單位'");
            SB.AppendFormat(" FROM [TK].dbo.COPTD,[TK].dbo.COPTC");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TD004 LIKE '4%'");
            SB.AppendFormat(" AND (TD008-TD009+TD024-TD025)>0");
            SB.AppendFormat(" AND TD021='Y' ");
            SB.AppendFormat(" AND TD016='N'");
            SB.AppendFormat(" AND TC001 IN ('A221', 'A222','A223','A228')");
            SB.AppendFormat(" GROUP BY TD005,TD004,TD006,TD010");
            SB.AppendFormat(" ORDER BY TD005,TD004,TD006,TD010");
            SB.AppendFormat(" ");

            return SB;

        }


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
            SETFASTREPORT2();
        }

        #endregion
    }
}
