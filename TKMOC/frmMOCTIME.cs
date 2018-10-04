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
    public partial class frmMOCTIME : Form
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

        public frmMOCTIME()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();

            SQL = SETSQL();

            Report report1 = new Report();
            report1.Load(@"REPORT\生產工時比較.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {


            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" SELECT CSTMB.MB001 AS '線代',CSTMB.MB002 AS '日期',CSTMB.MB003 AS '製令單',CSTMB.MB004 AS '製令',CSTMB.MB005 AS '總小時',CSTMB.MB007 AS '品號'");
            SB.AppendFormat(@" ,MOCTA.TA007 AS '單位',MOCTA.TA034 AS '品名',MOCTA.TA035 AS '規格',MOCTA.TA017 AS '生產量'");
            SB.AppendFormat(@" ,ISNULL([AVGTIME],0) AS '每個標準工時'");
            SB.AppendFormat(@" ,ISNULL([AVGTIME],0)*MOCTA.TA017 AS '標準總工時'");
            SB.AppendFormat(@" ,CSTMB.MB005*60 AS '實際總工時'");
            SB.AppendFormat(@" ,MD002 AS '線別'");
            SB.AppendFormat(@" ,(CSTMB.MB005*60-(ISNULL([AVGTIME],0)*MOCTA.TA017)) AS '工時差異'");
            SB.AppendFormat(@" FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD,[TK].dbo.CSTMB");
            SB.AppendFormat(@" LEFT JOIN [TKMOC].[dbo].[MOCCOSTTIME] ON [MOCCOSTTIME].[MB001]=[CSTMB].[MB007]");
            SB.AppendFormat(@" WHERE CSTMB.MB003=TA001 AND CSTMB.MB004=TA002");
            SB.AppendFormat(@" AND TA021=MD001");
            SB.AppendFormat(@" AND CSTMB.MB001 NOT IN ('08')  ");
            SB.AppendFormat(@" AND CSTMB.MB002>='{0}' AND CSTMB.MB002<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@" ORDER BY MD002,CSTMB.MB002,CSTMB.MB005");
            SB.AppendFormat(@" ");
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
