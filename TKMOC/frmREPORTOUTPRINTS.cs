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
        public frmREPORTOUTPRINTS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string SDAYS,string EDAYS)
        {
            StringBuilder SQL1 = new StringBuilder();
            string CHECK_TA021 = "";

            Report report1 = new Report();   
             
            report1.Load(@"REPORT\營貼\燒番麥方塊酥10g(品皇).frx");

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
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyy.MM.dd"), dateTimePicker2.Value.ToString("yyyy.MM.dd"));
        }
        #endregion
    }
}
