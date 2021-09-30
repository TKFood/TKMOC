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
using TKITDLL;

namespace TKMOC
{
    public partial class frmPORMOCIN : Form
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


        public frmPORMOCIN()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();

            SQL = SETSQL();

            Report report1 = new Report();
            report1.Load(@"REPORT\製令完工率.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
           

            StringBuilder SB = new StringBuilder();
 
            SB.AppendFormat(@" SELECT TA001 AS '製令單',TA002 AS '單號',TA003 AS '開工日',TA006 AS '品號',TA007 AS '單位'");
            SB.AppendFormat(@" ,TA015 AS '預計生產',TA017 AS '已生產',TA021 AS '線代',MD002 AS '線別',TA034 AS '品名',TA035 AS '規格'");
            SB.AppendFormat(@" ,ISNULL((SELECT SUM(TG011) FROM [TK].dbo.MOCTG WHERE TG004=TA006 AND TG014=TA001 AND TG015=TA002 ),0) AS '入庫量'");
            SB.AppendFormat(@" ,(ISNULL((SELECT SUM(TG011) FROM [TK].dbo.MOCTG WHERE TG004=TA006 AND TG014=TA001 AND TG015=TA002 ),0) /TA015) AS '完成率'");
            SB.AppendFormat(@" FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD");
            SB.AppendFormat(@" WHERE  TA021=MD001");
            SB.AppendFormat(@" AND TA003>='{0}' AND TA003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@" AND TA013='Y'");
            SB.AppendFormat(@" AND TA021<>'08'");
            SB.AppendFormat(@" ORDER BY TA003,TA021 DESC");
            SB.AppendFormat(@" ");
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
