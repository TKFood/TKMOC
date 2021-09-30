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
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKMOC
{
    public partial class frmREPORTMOCBOMMD : Form
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
        public frmREPORTMOCBOMMD()
        {
            InitializeComponent();
        }
        #region FUNCTION
        public void SETFASTREPORT(string MD003)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3(MD003);
            Report report1 = new Report();
            report1.Load(@"REPORT\原料-使用量差表.frx");


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
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL3(string MD003)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  SELECT TA014 AS '實際完工日期',TA001 AS '製令單別',TA002 AS '製令編號',TA006 AS '產品品號',TA034 AS '產品品名',TB003 AS '材料品號',TB012 AS '材料品名',TB004 AS '需領量',TB005 AS '實際領量',TB007 AS '領用單位',(TB005-TB004) AS '領用差異',(TB005-TB004)/ISNULL(NULLIF(TB005, 0),1) AS '實際損耗率',0 AS '標準損耗率',CONVERT(DECIMAL(12,2),TB005/MD006/MD007) AS '生產桶數',TA015 AS '預計產量',TA017 AS '已生產量',TA007 AS '產品單位'");
            SB.AppendFormat(@"  FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB,[TK].dbo.BOMMD");
            SB.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            SB.AppendFormat(@"  AND MD001=TA006");
            SB.AppendFormat(@"  AND MD003=TB003");
            SB.AppendFormat(@"  AND TA014>='{0}' AND TA009<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@"  AND TB003='{0}'",MD003);
            SB.AppendFormat(@"  ORDER BY TA014,TA001,TA002");
            SB.AppendFormat(@"  ");


            return SB;

        }
        #endregion

        #region BUTTON
        private void button15_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox1.Text.Trim()))
            {
                SETFASTREPORT(textBox1.Text.Trim());
            }
            else
            {
                MessageBox.Show("未輸入料號");
            }
           
        }
        #endregion
    }
}
