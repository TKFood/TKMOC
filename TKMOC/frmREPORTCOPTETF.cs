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
    public partial class frmREPORTCOPTETF : Form
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

        public frmREPORTCOPTETF()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1();
            Report report1 = new Report();
            report1.Load(@"REPORT\訂單變更明細表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            if(comboBox1.Text.Equals("Y"))
            {
                SB.AppendFormat(@" 
                            SELECT TE004 AS '變更日',TE006 AS '單頭變更原因',TF032 AS '單身變更原因',TE007 AS '客代',MA002 AS '客戶',TF001 AS '訂單單別',TF002 AS '訂單單號',TF004 AS '訂單序號',TF003 AS '訂單版次',TF005 AS '品號',TF006 AS '品名',TF015 AS '預交日',TF009 AS '數量',TF010 AS '單位'
                            FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPTE,[TK].dbo.COPTF,[TK].dbo.COPMA
                            WHERE TC001=TD001 AND TC002=TD002
                            AND TC001=TE001 AND TC002=TE002
                            AND TD001=TF001 AND TD002=TF002 AND TD003=TF104
                            AND TE007=MA001
                            AND COPTD.UDF01='Y'
                            AND TE004>='{0}' AND TE004<='{1}'
                            ORDER BY TE004,TF001,TF002,TF004,TF003
                            ",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            }
            else
            {
                SB.AppendFormat(@" 
                            SELECT TE004 AS '變更日',TE006 AS '單頭變更原因',TF032 AS '單身變更原因',TE007 AS '客代',MA002 AS '客戶',TF001 AS '訂單單別',TF002 AS '訂單單號',TF004 AS '訂單序號',TF003 AS '訂單版次',TF005 AS '品號',TF006 AS '品名',TF015 AS '預交日',TF009 AS '數量',TF010 AS '單位'
                            FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPTE,[TK].dbo.COPTF,[TK].dbo.COPMA
                            WHERE TC001=TD001 AND TC002=TD002
                            AND TC001=TE001 AND TC002=TE002
                            AND TD001=TF001 AND TD002=TF002 AND TD003=TF104
                            AND TE007=MA001
                            AND COPTD.UDF01<>'Y'
                            AND TE004>='{0}' AND TE004<='{1}'
                            ORDER BY TE004,TF001,TF002,TF004,TF003
                            ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            }
           

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
