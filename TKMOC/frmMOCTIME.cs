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
    public partial class frmMOCTIME : Form
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
        int result;
        SqlTransaction tran;

        public frmMOCTIME()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();
            StringBuilder SQL3 = new StringBuilder();

            SQL = SETSQL();
            SQL2 = SETSQL2();
            SQL3 = SETSQL3();

            Report report1 = new Report();
            report1.Load(@"REPORT\工時包裝.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2.ToString();
            TableDataSource table2 = report1.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SQL3.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

           
            SB.AppendFormat(@" SELECT TA003 AS '生產日',SUM(ROUND(TA015/INVMB.UDF10,2)) AS '預計總工時',SUM(ROUND(TA017/INVMB.UDF10,2)) AS '實際總工時'");
            SB.AppendFormat(@" FROM [TK].dbo.MOCTA,[TK].dbo.INVMB");
            SB.AppendFormat(@" WHERE TA006=MB001");
            SB.AppendFormat(@" AND TA021='09'");
            SB.AppendFormat(@" AND TA003 LIKE '{0}%'",dateTimePicker1.Value.ToString("yyyyMM"));
            SB.AppendFormat(@" AND INVMB.UDF10>0");
            SB.AppendFormat(@" GROUP BY TA003");
            SB.AppendFormat(@" ORDER BY TA003");
            SB.AppendFormat(@" ");
            SB.AppendFormat(@" ");


            return SB;

        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" SELECT TA003 AS '生產日',TA006 AS '品號',INVMB.MB002 AS '品名',TA007 AS '單位', TA015 AS '預計產量', TA017 AS '已生產量',INVMB.UDF10 AS '平均生產量/小時',ROUND(TA015/INVMB.UDF10,2) AS '預計總工時',ROUND(TA017/INVMB.UDF10,2) AS '實際總工時',TA001 AS '製令單',TA002 AS '單號'");
            SB.AppendFormat(@" FROM [TK].dbo.MOCTA,[TK].dbo.INVMB");
            SB.AppendFormat(@" WHERE TA006=MB001");
            SB.AppendFormat(@" AND TA021='09'");
            SB.AppendFormat(@" AND TA003 LIKE '{0}%'",dateTimePicker1.Value.ToString("yyyyMM"));
            SB.AppendFormat(@" AND INVMB.UDF10>0");
            SB.AppendFormat(@" ORDER BY TA003,TA006,INVMB.MB002,TA007,INVMB.UDF10 ");
            SB.AppendFormat(@" ");
            SB.AppendFormat(@" ");
            SB.AppendFormat(@" ");
            SB.AppendFormat(@" ");
            SB.AppendFormat(@" ");


            return SB;

        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" SELECT TA006 AS '品號',INVMB.MB002 AS '品名',TA003 AS '生產日',TA007 AS '單位',TA001 AS '製令單',TA002 AS '單號',TA015 AS '預計產量',TA017 AS '已生產量',INVMB.UDF10 AS '平均生產量/小時'");
            SB.AppendFormat(@" FROM [TK].dbo.MOCTA,[TK].dbo.INVMB");
            SB.AppendFormat(@" WHERE TA006=MB001");
            SB.AppendFormat(@" AND TA021='09'");
            SB.AppendFormat(@" AND TA003 LIKE '201908%'");
            SB.AppendFormat(@" AND INVMB.UDF10=0");
            SB.AppendFormat(@" ORDER BY TA003,TA006");
            SB.AppendFormat(@"  ");


            return SB;

        }

        public void UPDATEINVMBUDF10()
        {
            DateTime dt = DateTime.Now;
            DateTime lastmoneth = dt.AddMonths(-1);

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

             
                sbSql.AppendFormat("  UPDATE [TK].dbo.INVMB");
                sbSql.AppendFormat("  SET UDF10=平均生產量小時");
                sbSql.AppendFormat("  FROM (SELECT SUM(CSTMB.MB005) AS '人時',SUM(TA017) AS '生產量',ROUND(SUM(TA017)/SUM(CSTMB.MB005),4) AS '平均生產量小時',TA006,TA007,INVMB.MB002");
                sbSql.AppendFormat("  FROM [TK].dbo.CSTMB,[TK].dbo.MOCTA,[TK].dbo.INVMB");
                sbSql.AppendFormat("  WHERE CSTMB.MB003=TA001 AND CSTMB.MB004=TA002");
                sbSql.AppendFormat("  AND INVMB.MB001=TA006");
                sbSql.AppendFormat("  AND CSTMB.MB001='09'");
                sbSql.AppendFormat("  AND (CSTMB.MB002 LIKE '{0}%' OR CSTMB.MB002 LIKE '{1}%')",dt.ToString("yyyyMM"), lastmoneth.ToString("yyyyMM"));
                sbSql.AppendFormat("  GROUP BY TA006,TA007,INVMB.MB002) AS TEMP");
                sbSql.AppendFormat("  WHERE TEMP.TA006=MB001 ");
                sbSql.AppendFormat("   ");
                sbSql.AppendFormat("   ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消

                    MessageBox.Show("失敗");
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("成功");
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            UPDATEINVMBUDF10();
        }

        #endregion


    }
}
