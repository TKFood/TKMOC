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
    public partial class frmREPORTCHECKPURTDPURTH : Form
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

        public frmREPORTCHECKPURTDPURTH()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string SDAYS,string EDAYS)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1.AppendFormat(@"
                                SELECT TD012 AS '預計進貨日'
                                ,TC004 AS '廠商'
                                ,MA002 AS '廠商名'
                                ,TC001 AS '採購單別'
                                ,TC002 AS '採購單號'
                                ,TD003 AS '採購序號'
                                ,TD004 AS '品號'
                                ,TD005 AS '品名'
                                ,TD008 AS '採購數量'
                                ,TD009 AS '單位'
                                ,ISNULL(SUMTH007,0) AS '已進貨數量'
                                FROM [TK].dbo.PURMA,[TK].dbo.PURTC,[TK].dbo.PURTD
                                LEFT JOIN (SELECT SUM(TH007) SUMTH007,TH011,TH012,TH013 FROM [TK].dbo.PURTH GROUP BY TH011,TH012,TH013) AS TEMP  ON TH011=TD001 AND TH012=TD002 AND TH013=TD003
                                WHERE TC001=TD001 AND TC002=TD002
                                AND TD018='Y'
                                AND MA001=TC004
                                AND (TD008>ISNULL(SUMTH007,0))
                                AND TD012>='{0}' AND TD012<='{1}'

                                ORDER BY TD012,TC004,TD004
 
                                ", SDAYS,EDAYS);

            Report report1 = new Report();
            report1.Load(@"REPORT\採購單是否有進貨.frx");

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


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        #endregion
    }
}
