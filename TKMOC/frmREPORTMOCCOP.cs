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
    public partial class frmREPORTMOCCOP : Form
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

        public frmREPORTMOCCOP()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1();
            Report report1 = new Report();
            report1.Load(@"REPORT\製令準時完工率數量達交率.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT TC053 AS '客戶',TA001 AS '製令單',TA002 AS '製令編號',TA006 AS '品號',TA034 AS '品名',TA007 AS '生產單位',TA009 AS '預計開工日',TA010 AS '預計完工日',TA014 AS '實際完工日',TA015 AS '預計產量',TA017 AS '已生產量',TA026 AS '訂單別',TA027 AS '訂單號',TA028 AS '訂單序',[OLDNUM] AS '訂單量'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TA017),0) FROM [TK].dbo.MOCTA A WHERE A.TA011 IN ('Y','y') AND A.TA026=[COPTD].TD001 AND A.TA027=[COPTD].TD002 AND A.TA028=[COPTD].TD003 AND A.TA006=[COPTD].TD004  ) AS '訂單總生產量'");
            SB.AppendFormat(" ,ISNULL(((SELECT ISNULL(SUM(TA017),0) FROM [TK].dbo.MOCTA A WHERE A.TA011 IN ('Y','y') AND A.TA026=[COPTD].TD001 AND A.TA027=[COPTD].TD002 AND A.TA028=[COPTD].TD003 AND A.TA006=[COPTD].TD004  ) -[OLDNUM]),0) AS '生產數量是否滿足訂單'");
            SB.AppendFormat(" ,[COPTD].TD013 AS '訂單預交日'");
            SB.AppendFormat(" ,DATEDIFF (DAY,[COPTD].TD013,TA014) AS '是否延遲訂單預交'");
            SB.AppendFormat(" ,DATEDIFF (DAY,TA010,TA014) AS '是否延遲製令完工'");
            SB.AppendFormat(" ,ISNULL((TA017-TA015),0) AS '製令生產數量生否>預計生產'");
            SB.AppendFormat(" FROM [TK].dbo.MOCTA");
            SB.AppendFormat(" LEFT JOIN [TK].[dbo].[VCOPTDINVMD] ON [VCOPTDINVMD].TD001=TA026 AND [VCOPTDINVMD].TD002=TA027 AND [VCOPTDINVMD].TD003=TA028");
            SB.AppendFormat(" LEFT JOIN [TK].[dbo].[COPTD] ON [COPTD].TD001=TA026 AND [COPTD].TD002=TA027 AND [COPTD].TD003=TA028");
            SB.AppendFormat(" LEFT JOIN [TK].[dbo].[COPTC] ON [COPTC].TC001=TA026 AND [COPTC].TC002=TA027 ");
            SB.AppendFormat(" WHERE TA001 IN ('A510','A511')");
            SB.AppendFormat(" AND TA006 LIKE '4%'");
            SB.AppendFormat(" AND TA009>='{0}' AND TA009<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ORDER BY TC053,TA006 ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

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
