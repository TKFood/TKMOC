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
    public partial class frmREPORTCOPBARS : Form
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

        public frmREPORTCOPBARS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();

            SQL = SETSQL();

            report1.Load(@"REPORT\訂單預計生產的桶數報表.frx");

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
            StringBuilder QUERY = new StringBuilder();

            if(comboBox1.Text.Equals("是"))
            {
                QUERY.AppendFormat(@" 
                                     AND COPTD.UDF01='Y'
                                    ");
            }
            else
            {
                QUERY.AppendFormat(@" 
                                   
                                    ");
            }

            SB.AppendFormat(@" 
                                SELECT TC001 AS '訂單',TC002 AS '單號',TD003 AS '序號',TC003 AS '訂單日期',TC004 AS '客戶代號'
                                ,MA002 AS '客戶',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',(TD008+TD024) AS '訂單數量'
                                ,MB068 AS '生產別',MC1.MC004  AS 'MC1MC004',MD1.MD003 AS 'MD1MD003',MD1.MD006 AS 'MD1MD006'
                                ,MD1.MD007 AS 'MD1MD007',MC2.MC001 AS 'MC2MC001',MC2.MC004  AS 'MC2MC004'
                                ,((TD008+TD024)/MC1.MC004*MD1.MD006*(1+MD1.MD007)/MC2.MC004)  AS 'BAR'
                                ,TD013 AS '預交日'
                                FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.INVMB,[TK].dbo.COPMA,[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1,[TK].dbo.BOMMC MC2
                                WHERE TC001=TD001 AND TC002=TD002
                                AND TD004=MB001
                                AND TC004=MA001
                                AND TD004=MC1.MC001
                                AND MC1.MC001=MD1.MD001
                                AND MC2.MC001=MD1.MD003
                                AND TC027='Y'
                                {2}
                                AND MD1.MD003 LIKE '301%'
                                AND MD1.MD003 NOT LIKE '30100002%'
                                AND MB068 IN ('09')
                                AND TC003>='{0}' AND TC003<='{1}'
                                UNION ALL
                                SELECT TC001,TC002,TD003,TC003,TC004,MA002,TD004,TD005,TD006,(TD008+TD024),MB068,MC1.MC004 MC1MC004,MD1.MD003,MD1.MD006,MD1.MD007,MC2.MC001,MC2.MC004 MC2MC004,((TD008+TD024)/MC1.MC004) AS 'BAR',TD013 AS '預交日'
                                FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.INVMB,[TK].dbo.COPMA,[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1,[TK].dbo.BOMMC MC2
                                WHERE TC001=TD001 AND TC002=TD002
                                AND TD004=MB001
                                AND TC004=MA001
                                AND TD004=MC1.MC001
                                AND MC1.MC001=MD1.MD001
                                AND MC2.MC001=MD1.MD003
                                AND TC027='Y'
                                   {2}
                                AND MD1.MD003 LIKE '301%'
                                AND MD1.MD003 NOT LIKE '30100002%'
                                AND MB068 IN ('02','03')
                                AND TC003>='{0}' AND TC003<='{1}' 

                            ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), QUERY.ToString());

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
