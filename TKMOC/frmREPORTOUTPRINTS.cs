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
        private bool isProcessing = false; // 宣告旗標

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
                                ,[BOXNAMES]
                                ,[ORDRES]
                                FROM [TKMOC].[dbo].[TBOUTBOXNAMES]
                                WHERE [ISCLOSED]='N'
                                ORDER BY [ORDRES]
                                    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("BOXNAMES", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "BOXNAMES";
            comboBox1.DisplayMember = "BOXNAMES";
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
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
        
        public DataTable FIND_TBOUTBOXNAMES(string BOXNAMES)
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
                                [ID], [BOXNAMES], [MB001], [ORDRES], [ISCLOSED]
                                FROM [TKMOC].[dbo].[TBOUTBOXNAMES]
                                WHERE [ISCLOSED]='N'
                                AND ([BOXNAMES] LIKE @BOXNAMES OR [MB001] LIKE @MB001)
                            ");

            using (SqlCommand cmd = new SqlCommand(Sequel.ToString(), sqlConn))
            {
                // 建立參數並將值設定為 %輸入值%
                cmd.Parameters.AddWithValue("@BOXNAMES", $"%{BOXNAMES}%");
                cmd.Parameters.AddWithValue("@MB001", $"%{BOXNAMES}%");

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sqlConn.Open();
                da.Fill(dt);
                sqlConn.Close();

                if (dt != null && dt.Rows.Count >= 1)
                {
                    return dt;
                }
                else
                {
                    return null;
                }
            }

        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 檢查旗標，避免重複執行
            if (isProcessing) return;

            // 設定旗標，表示正在處理中
            isProcessing = true;
            try
            {
                textBox1.Text = null;
                // 確保 comboBox1.Text 不為空再查詢
                if (string.IsNullOrWhiteSpace(comboBox1.Text)) return;

                DataTable DT = FIND_TBOUTBOXNAMES(comboBox1.Text.ToString());
                if (DT != null && DT.Rows.Count >= 1)
                {
                    // 此行會觸發 textBox1_TextChanged，但因為 isProcessing=true 而被阻擋
                    textBox1.Text = DT.Rows[0]["MB001"].ToString();
                }
            }
            finally
            {
                // 處理完成，重設旗標
                isProcessing = false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // 檢查旗標，避免重複執行
            if (isProcessing) return;

            // 設定旗標，表示正在處理中
            isProcessing = true;
            try
            {
                string MB001 = textBox1.Text.Trim();
                if (string.IsNullOrWhiteSpace(MB001)) return;

                DataTable DT = FIND_TBOUTBOXNAMES(MB001);
                if (DT != null && DT.Rows.Count >= 1)
                {
                    // 此行會觸發 comboBox1_SelectedIndexChanged，但因為 isProcessing=true 而被阻擋
                    comboBox1.Text = DT.Rows[0]["BOXNAMES"].ToString();
                }
            }
            finally
            {
                // 處理完成，重設旗標
                isProcessing = false;
            }
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {             
            SETFASTREPORT(comboBox1.Text.ToString(),dateTimePicker1.Value.ToString("yyyy.MM.dd"), dateTimePicker2.Value.ToString("yyyy.MM.dd"));
            SETFASTREPORT(comboBox1.Text.ToString(),dateTimePicker1.Value.ToString("yyyy.MM.dd"), dateTimePicker2.Value.ToString("yyyy.MM.dd"));
        }



        #endregion

        
    }
}
