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
    public partial class frmREPORTCOPMOC : Form
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
        Report report1 = new Report();

        public frmREPORTCOPMOC()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();

            SQL = SETSQL();

            report1.Load(@"REPORT\查訂單-製令-入庫.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            

            StringBuilder SB = new StringBuilder();
          
            SB.AppendFormat(@" SELECT COPTD.TD013 AS '預交日',COPTD.TD001 AS '訂單', COPTD.TD002 AS '單號', COPTD.TD003 AS '序號', COPTD.TD004 AS '品號', COPTD.TD005 AS '品名', COPTD.TD008 AS '下訂量', COPTD.TD009 AS '已出貨', COPTD.TD010 AS '單位'");
            SB.AppendFormat(@" ,MOCTA.TA001 AS '製令',MOCTA.TA002 AS '製令號',MOCTA.TA009 AS '生產日',MOCTA.TA017 AS '生產量'");
            SB.AppendFormat(@" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(@" LEFT JOIN [TK].dbo.MOCTA ON MOCTA.TA026=COPTD.TD001 AND MOCTA.TA027=COPTD.TD002");
            SB.AppendFormat(@" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(@" AND COPTD.TD013>='{0}' AND COPTD.TD013<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@" AND COPTD.TD008>0");
            SB.AppendFormat(@" ORDER BY COPTD.TD013,COPTD.TD001,COPTD.TD004");
            SB.AppendFormat(@" ");

            return SB;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
