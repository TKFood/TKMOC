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
    public partial class frmREPORTOUTPRINTS : Form
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

        int LIMITSMONTHS = 0;

        public frmREPORTOUTPRINTS()
        {
            InitializeComponent();

            comboBox1load();
            comboBox1load2();
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
                                SELECT 
                                [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKMOC].[dbo].[TBPARA]
                                WHERE [KIND]='外標列印'
                                ORDER BY CONVERT(INT,[PARANAME])
                                    ");
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

        public void comboBox1load2()
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
                                SELECT 
                                [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKMOC].[dbo].[TBPARA]
                                WHERE [KIND]='有效期限'
                               
                                    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "PARANAME";
            comboBox2.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try 
            {
                LIMITSMONTHS = Convert.ToInt32(comboBox2.SelectedValue.ToString());

                if (LIMITSMONTHS > 0)
                {
                    DateTime CAL_DAYS = dateTimePicker1.Value;
                    CAL_DAYS = CAL_DAYS.AddDays(-1);
                    CAL_DAYS = CAL_DAYS.AddMonths(LIMITSMONTHS);

                    dateTimePicker2.Value = CAL_DAYS;
                }
                //MessageBox.Show(LIMITSMONTHS.ToString());
            }
            catch
            {

            }
            
        }
        public void SETFASTREPORT(string REPORTS,string SDAYS,string EDAYS)
        {
            StringBuilder SQL1 = new StringBuilder();
            string CHECK_TA021 = "";
               
            Report report1 = new Report();

            string REPORTS_DIR = @"REPORT\營貼\"+ REPORTS + ".frx";

            report1.Load(REPORTS_DIR);

            //SQL1 = SETSQL1(CHECK_TA021, SDAY, EDAY);


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);

            //report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            //TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            //table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", SDAYS);
            report1.SetParameterValue("P2", EDAYS);

            report1.Preview = previewControl1;
            report1.Show();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {           
            SETFASTREPORT(comboBox1.Text.ToString(),dateTimePicker1.Value.ToString("yyyy.MM.dd"), dateTimePicker2.Value.ToString("yyyy.MM.dd"));
        }
        #endregion

     
    }
}
