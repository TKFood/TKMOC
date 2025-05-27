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
    public partial class frmREPORTDAILY1: Form
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

        public frmREPORTDAILY1()
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
            Sequel.AppendFormat(@"
                                    SELECT LTRIM(RTRIM(MD001)) MD001,MD002
                                    FROM [TK].dbo.CMSMD
                                    WHERE MD001 IN ('02','03','04','09','08','12')");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);

            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD001";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void SETFASTREPORT(string TA021,string TA003)
        {
            StringBuilder SQL1 = new StringBuilder();
            Report report1 = new Report(); 

            if (TA021.Equals("09"))
            {
                SQL1 = SETSQL1(TA021, TA003);
                report1.Load(@"REPORT\外包課包裝報表.frx");
            }
            else if(TA021.Equals("04"))
            {
                SQL1 = SETSQL2(TA021, TA003);
                report1.Load(@"REPORT\生產產能報表-手工.frx");
            }  
            else if (TA021.Equals("02"))
            {
                SQL1 = SETSQL2(TA021, TA003);
                report1.Load(@"REPORT\生產產能報表.frx");
            }
            else if (TA021.Equals("08"))
            {
                SQL1 = SETSQL2(TA021, TA003);
                report1.Load(@"REPORT\生產產能報表-烘培生產.frx");
            }
            else if (TA021.Equals("12"))
            {
                SQL1 = SETSQL1(TA021, TA003);
                report1.Load(@"REPORT\生產產能報表-烘培外包.frx");
            }





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

        public StringBuilder SETSQL1(string TA021, string TA003)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                              SELECT 
                            YEAR(CONVERT(DATETIME,TA003) ) AS 'YEARS'
                            ,MONTH(CONVERT(DATETIME,TA003) ) AS 'MONTHS'
                            ,DAY(CONVERT(DATETIME,TA003) ) AS 'DAYS'
                            ,TA001+TA002 AS '單號'
                            ,TA006 AS '生產品號'
                            ,TA034 AS '生產品名'
                            ,TB003 AS '領料品號'
                            ,TB012 AS '領料品名'
                            FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB
                            WHERE TA001=TB001 AND TA002=TB002
                            AND (TB003 LIKE '3%' OR TB003 LIKE '4%')
                            AND TA021='{0}'
                            AND TA002 LIKE '%{1}%'
       
                            ORDER BY TA001,TA002


                            ", TA021, TA003);

            return SB;


        }

        public StringBuilder SETSQL2(string TA021, string TA003)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT 
                            YEAR(CONVERT(DATETIME,TA003) ) AS 'YEARS'
                            ,MONTH(CONVERT(DATETIME,TA003) ) AS 'MONTHS'
                            ,DAY(CONVERT(DATETIME,TA003) ) AS 'DAYS'
                            ,*
                            FROM [TK].dbo.MOCTA
                            WHERE TA021='{0}'
                            AND TA002 LIKE '%{1}%'
       
                            ORDER BY TA001,TA002
                           


                            ", TA021, TA003);

            return SB;

        }

        #endregion

        #region BUTTON
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox1.SelectedValue.ToString(),dateTimePicker1.Value.ToString("yyyyMMdd")); 
        }

        #endregion
    }
}
