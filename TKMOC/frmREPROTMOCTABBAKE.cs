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
    public partial class frmREPROTMOCTABBAKE : Form
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

        public frmREPROTMOCTABBAKE()
        {
            InitializeComponent();

            comboBox1load();

        }

        #region FUNCTION
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [KIND],[PARAID],[PARANAME] FROM [TKMOC].[dbo].[TBPARA] WHERE [KIND]='BAKE'  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARAID";
            comboBox1.DisplayMember = "PARAID";
            sqlConn.Close();


        }


        public void SETFASTREPORT(string TA001,string TA003)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1(TA001, TA003);

            Report report1 = new Report();
            report1.Load(@"REPORT\原物料添加表-烘焙.frx");

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

        public StringBuilder SETSQL1(string TA001, string TA003)
        {
            StringBuilder SB = new StringBuilder();

           
            SB.AppendFormat(@" 
                            SELECT 
                            TA001 AS '製令',TA002 AS '製令單',TA003 AS '生產日',TA006 AS '生產品號',MB1.MB002 AS '生產品名',TA015 AS '生產量',TA007 AS '生產單位'
                            ,TB003 AS '原/物料品號',MB2.MB002 AS '原/物料品名',TB004 AS '需領料數量',TB007 AS '領料單位'
                            ,(YEAR(TA003)-1911) AS 'YEARS',MONTH(TA003) AS 'MONTHS',DAY(TA003) AS 'DAYS'
                            FROM [TK].dbo.MOCTA
                            LEFT JOIN [TK].dbo.INVMB MB1 ON MB1.MB001=TA006
                            ,[TK].dbo.MOCTB
                            LEFT JOIN [TK].dbo.INVMB MB2 ON MB2.MB001=TB003
                            WHERE TA001=TB001 AND TA002=TB002
                            AND TA001='{0}'
                            AND TA003='{1}'
                            ORDER BY TA001,TA002,TA006,TB003

                            ", TA001, TA003);

            return SB;

        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox1.Text.ToString(),dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        #endregion
    }
}
