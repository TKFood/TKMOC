using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Globalization;
using Calendar.NET;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmREPORTCOSTCO : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();


        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        int rownum = 0;
        DataSet ds1 = new DataSet();

        public Report report1 { get; private set; }

        public frmREPORTCOSTCO()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\COSTCO-領料表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;

            report1.Preview = previewControl1;
            report1.Show();
        }


        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();

            //,CASE WHEN TA006 NOT LIKE '4%' THEN CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)) ELSE 0 END AS '桶數'
            //,CASE WHEN TA006 LIKE '4%' THEN CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1)) ELSE 0 END AS '箱數'

            FASTSQL.AppendFormat(@"    
                                SELECT TC003 AS '領料日期',TC001+TC002 AS '領料單號',TC014 AS '單據日期',TE004 AS '材料品號',TE005 AS '領料數量',TE006 AS '單位',MC002 AS '庫別',TE009 AS '製程代號',TE010 AS '批號'
                                ,TE011+TE012 AS '製令單號',TE013 AS '領料說明',TE014 AS '備註'
                                ,TE017 AS '品名',TE018 AS '規格'
                                FROM [TK].dbo.MOCTC,[TK].dbo.MOCTE
                                LEFT JOIN [TK].dbo.CMSMC ON MC001=TE008
                                WHERE TC001=TE001 AND TC002=TE002
                                AND TC001 LIKE 'A54%'
                                AND TE011+TE012 IN (SELECT TA001+TA002 FROM [TKMOC].dbo.COSTCO)
                                ORDER BY TE011+TE012
                                ");

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2()
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\COSTCO-退料表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL2();
            Table.SelectCommand = SQL;

            report1.Preview = previewControl2;
            report1.Show();
        }


        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();

            //,CASE WHEN TA006 NOT LIKE '4%' THEN CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)) ELSE 0 END AS '桶數'
            //,CASE WHEN TA006 LIKE '4%' THEN CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1)) ELSE 0 END AS '箱數'

            FASTSQL.AppendFormat(@"    
                                SELECT TC003 AS '退料日期',TC001+TC002 AS '退料單號',TC014 AS '單據日期',TE004 AS '材料品號',TE005 AS '退料數量',TE006 AS '單位',MC002 AS '庫別',TE009 AS '製程代號',TE010 AS '批號'
                                ,TE011+TE012 AS '製令單號',TE013 AS '領料說明',TE014 AS '備註'
                                ,TE017 AS '品名',TE018 AS '規格'
                                FROM [TK].dbo.MOCTC,[TK].dbo.MOCTE
                                LEFT JOIN [TK].dbo.CMSMC ON MC001=TE008
                                WHERE TC001=TE001 AND TC002=TE002
                                AND TC001 LIKE 'A56%'
                                AND TE011+TE012 IN (SELECT TA001+TA002 FROM [TKMOC].dbo.COSTCO)
                                ORDER BY TE011+TE012
                                ");

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT3()
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\COSTCO-入庫表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL3();
            Table.SelectCommand = SQL;

            report1.Preview = previewControl3;
            report1.Show();
        }


        public string SETFASETSQL3()
        {
            StringBuilder FASTSQL = new StringBuilder();

            //,CASE WHEN TA006 NOT LIKE '4%' THEN CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)) ELSE 0 END AS '桶數'
            //,CASE WHEN TA006 LIKE '4%' THEN CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1)) ELSE 0 END AS '箱數'

            FASTSQL.AppendFormat(@"    
                                SELECT TF003 AS '入庫日期',TF001+TF002 AS '單別-單號',TF012  AS '單據日期',TG004  AS '品號',TG005  AS '品名',TG006  AS '規格',TG011  AS '入庫數量',TG007  AS '單位',TG013 AS '驗收數量','合格' AS '檢驗狀態',TG014+TG015  AS '製令編號',TG017 AS '批號',TG020 AS '備註',MC002 AS '庫別'
                                FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG
                                LEFT JOIN [TK].dbo.CMSMC ON MC001=TG010
                                WHERE TF001=TG001 AND TF002=TG002
                                AND TG014+TG015 IN (SELECT TA001+TA002 FROM [TKMOC].dbo.COSTCO)
                                ORDER BY TG014+TG015
                                ");

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3();
        }

        #endregion


    }
}
