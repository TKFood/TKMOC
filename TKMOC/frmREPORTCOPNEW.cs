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
    public partial class frmREPORTCOPNEW : Form
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

        public frmREPORTCOPNEW()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();

            SQL = SETSQL();

            Report report1 = new Report();
            if (comboBox3.Text.Equals("明細"))
            {
                report1.Load(@"REPORT\訂單統計表.frx");
            }
            else if (comboBox3.Text.Equals("月報"))
            {
                report1.Load(@"REPORT\訂單統計表-月報.frx");
            }

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder TD001 = new StringBuilder();
            if (checkBox1.Checked == true)
            {
                TD001.Append("'A221',");
            }
            if (checkBox2.Checked == true)
            {
                TD001.Append("'A222',");
            }

            if (checkBox4.Checked == true)
            {
                TD001.Append("'A225',");
            }
            if (checkBox5.Checked == true)
            {
                TD001.Append("'A226',");
            }
            if (checkBox6.Checked == true)
            {
                TD001.Append("'A227',");
            }
            if (checkBox7.Checked == true)
            {
                TD001.Append("'A223',");
            }
            if (checkBox3.Checked == true)
            {
                TD001.Append("'A228',");
            }
            TD001.Append("''");

            StringBuilder TC027 = new StringBuilder();
            if (comboBox1.Text.ToString().Equals("已確認"))
            {
                TC027.Append(" 'Y',");
            }
            else if (comboBox1.Text.ToString().Equals("未確認"))
            {
                TC027.Append(" 'N',");
            }
            else if (comboBox1.Text.ToString().Equals("全部"))
            {
                TC027.Append(" 'Y','N', ");
            }
            TC027.Append("''");

            StringBuilder ORDER = new StringBuilder();
            if (comboBox2.Text.ToString().Equals("依品名排序"))
            {
                ORDER.Append(" ORDER BY 品名,規格,日期,單位,客戶");
            }
            else if (comboBox2.Text.ToString().Equals("依日期排序"))
            {
                ORDER.Append(" ORDER BY 日期,品名,規格,單位,客戶");
            }

            StringBuilder SB = new StringBuilder();
            SB.AppendFormat(@" SELECT 品名,規格,日期,SUM(訂單數量) AS '訂單數量',SUM(訂單未交量) AS '訂單未交量',單位,客戶");
            SB.AppendFormat(@" FROM (");
            SB.AppendFormat(@" SELECT   TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TC053  AS '客戶' ,TD013 AS '日期',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
            SB.AppendFormat(@" ,(CASE WHEN MB004=TD010 THEN (TD008-TD009) ELSE (TD008-TD009)*MD004 END) AS '訂單數量'");
            SB.AppendFormat(@" ,MB004 AS '單位'");
            SB.AppendFormat(@" ,(TD008-TD009) AS '訂單未交量'");
            SB.AppendFormat(@" ,TD010 AS '訂單單位' ");
            SB.AppendFormat(@" ,(CASE WHEN ISNULL(MD002,'')<>'' THEN MD002 ELSE TD010 END ) AS '換算單位'");
            SB.AppendFormat(@" ,(CASE WHEN MD003>0 THEN MD003 ELSE 1 END) AS '分子'");
            SB.AppendFormat(@" ,(CASE WHEN MD004>0 THEN MD004 ELSE (TD008-TD009) END ) AS '分母'");
            SB.AppendFormat(@" FROM [TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(@" LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
            SB.AppendFormat(@" WHERE TD004=MB001");
            SB.AppendFormat(@" AND TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(@" AND TD004 LIKE '4%'");
            SB.AppendFormat(@" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@"  AND TC001 IN ({0}) ", TD001.ToString());
            SB.AppendFormat(@" AND (TD008-TD009)>0  ");
            SB.AppendFormat(@" AND TC027 IN ({0})  )", TC027.ToString());
            SB.AppendFormat(@" AS TEMP");
            SB.AppendFormat(@" GROUP BY 品名,規格,日期,單位,客戶");
            SB.AppendFormat(@" {0}", ORDER.ToString());
            SB.AppendFormat("");



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
