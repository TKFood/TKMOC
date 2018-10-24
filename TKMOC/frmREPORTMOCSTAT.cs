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
    public partial class frmREPORTMOCSTAT : Form
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

        public frmREPORTMOCSTAT()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1();
            SQL2 = SETSQL2();

            Report report1 = new Report();
            report1.Load(@"REPORT\生產日報表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();
           
            SB.AppendFormat(" SELECT TA001 AS '製令單',TA002 AS '單號',TA003 AS '開工日',TA006 AS '品號',TA007 AS '單位'");
            SB.AppendFormat(" ,TA015 AS '預計生產',TA017 AS '已生產',TA021 AS '線代',MD002 AS '線別',TA034 AS '品名',TA035 AS '規格'");
            SB.AppendFormat(" ,ISNULL((SELECT SUM(TG011) FROM [TK].dbo.MOCTG WHERE TG004=TA006 AND TG014=TA001 AND TG015=TA002 ),0) AS '入庫量'");
            SB.AppendFormat(" ,(ISNULL((SELECT SUM(TG011) FROM [TK].dbo.MOCTG WHERE TG004=TA006 AND TG014=TA001 AND TG015=TA002 ),0) /TA015) AS '完成率'");
            SB.AppendFormat(" FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD");
            SB.AppendFormat(" WHERE  TA021=MD001");
            SB.AppendFormat(" AND TA003>='{0}' AND TA003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TA013='Y'");
            SB.AppendFormat(" AND TA021<>'08'");
            SB.AppendFormat(" ORDER BY TA003,TA021 DESC");
            SB.AppendFormat(" ");

            return SB;

        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB2 = new StringBuilder();

            SB2.AppendFormat(" SELECT CSTMB.MB001 AS '線代',CSTMB.MB002 AS '日期',CSTMB.MB003 AS '製令單',CSTMB.MB004 AS '製令',CSTMB.MB005 AS '總小時',CSTMB.MB007 AS '品號'");
            SB2.AppendFormat(" ,MOCTA.TA007 AS '單位',MOCTA.TA034 AS '品名',MOCTA.TA035 AS '規格',MOCTA.TA017 AS '生產量'");
            SB2.AppendFormat(" ,ISNULL([AVGTIME],0) AS '每個標準工時'");
            SB2.AppendFormat(" ,ISNULL([AVGTIME],0)*MOCTA.TA017 AS '標準總工時'");
            SB2.AppendFormat(" ,CSTMB.MB005*60 AS '實際總工時'");
            SB2.AppendFormat(" ,MD002 AS '線別'");
            SB2.AppendFormat(" ,(CSTMB.MB005*60-(ISNULL([AVGTIME],0)*MOCTA.TA017)) AS '工時差異'");
            SB2.AppendFormat(" FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD,[TK].dbo.CSTMB");
            SB2.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[MOCCOSTTIME] ON [MOCCOSTTIME].[MB001]=[CSTMB].[MB007]");
            SB2.AppendFormat(" WHERE CSTMB.MB003=TA001 AND CSTMB.MB004=TA002");
            SB2.AppendFormat(" AND TA021=MD001");
            SB2.AppendFormat(" AND CSTMB.MB001 NOT IN ('08')");
            SB2.AppendFormat(" AND CSTMB.MB002>='{0}' AND CSTMB.MB002<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB2.AppendFormat(" ORDER BY MD002,CSTMB.MB002,CSTMB.MB005");
            SB2.AppendFormat(" ");
            SB2.AppendFormat(" ");

            return SB2;

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
