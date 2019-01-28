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
    public partial class frmRECOPSTILL : Form
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

        public frmRECOPSTILL()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1();
            Report report1 = new Report();
            report1.Load(@"REPORT\訂單未出完且未結案報表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT TC053 AS '客戶',TD013 AS '預交日',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD008 AS '訂單量',TD009 AS '出貨量',TD024 AS '贈品量',TD025 AS '贈品已交量',(TD008-TD009+TD024-TD025) AS '總未出貨量',TD010 AS '單位',TD001 AS '訂單',TD002 AS '單號',TD003 AS '序號',TA001 AS '製令',TA002 AS '製令單',TA009 AS '預計開工',TA015 AS '預計產量',TA017 AS '已生產量' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD ");
            SB.AppendFormat(" LEFT JOIN [TK].dbo.MOCTA ON TA026=TD001 AND TA027=TD002 AND TA028=TD003");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TD004 LIKE '4%'");
            SB.AppendFormat(" AND TD004 NOT LIKE '410%'");
            SB.AppendFormat(" AND (TD008-TD009+TD024-TD025)>0");
            SB.AppendFormat(" AND TD021='Y'");
            SB.AppendFormat(" AND TD016='N'");
            SB.AppendFormat(" AND TC001 IN ('A221', 'A222','A223','A228')");
            SB.AppendFormat(" ORDER BY TC001,TC053,TD013,TD004");
            SB.AppendFormat(" ");



            return SB;

        }
        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2();
            Report report2 = new Report();
            report2.Load(@"REPORT\訂單未出完且未結案的品號報表.frx");

            report2.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report2.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report2.Preview = previewControl2;
            report2.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',SUM(TD008) AS '訂單量',SUM(TD009)  AS '出貨量',SUM(TD024)  AS '贈品量',SUM(TD025)  AS '贈品已交量',SUM((TD008-TD009+TD024-TD025)) AS '總未出貨量',TD010 AS '單位'");
            SB.AppendFormat(" FROM [TK].dbo.COPTD,[TK].dbo.COPTC");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TD004 LIKE '4%'");
            SB.AppendFormat(" AND TD004 NOT LIKE '410%'");
            SB.AppendFormat(" AND (TD008-TD009+TD024-TD025)>0");
            SB.AppendFormat(" AND TD021='Y' ");
            SB.AppendFormat(" AND TD016='N'");
            SB.AppendFormat(" AND TC001 IN ('A221', 'A222','A223','A228')");
            SB.AppendFormat(" GROUP BY TD005,TD004,TD006,TD010");
            SB.AppendFormat(" ORDER BY TD005,TD004,TD006,TD010");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT3()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3();
            Report report3 = new Report();
            report3.Load(@"REPORT\製令明細表.frx");

            report3.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report3.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report3.Preview = previewControl3;
            report3.Show();
        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT MD002 AS '線別',TA009 AS '預計開工',TA034 AS '品名',TA001 AS '製令',TA002 AS '製令單',TA015 AS '預計產量',TA017 AS '已生產量',(TA015-TA017) AS '未生產量',TA007 AS '單位'");
            SB.AppendFormat(" FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD");
            SB.AppendFormat(" WHERE TA021=MD001");
            SB.AppendFormat(" AND TA009>='{0}' AND TA009<='{1}'",dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TA021 IN ('02','03','04','05','09','10')");
            SB.AppendFormat(" ORDER BY TA021,TA002,TA034");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT4()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL4();
            Report report4 = new Report();
            report4.Load(@"REPORT\訂單排產狀況表.frx");

            report4.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report4.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report4.Preview = previewControl4;
            report4.Show();
        }

        public StringBuilder SETSQL4()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT TC053 AS '客戶',TD013 AS '預計交貨日',TD004 AS '訂單品號',TD005 AS '訂單品名',TD006 AS '規格',TD008 AS '訂單量',TD009 AS '出貨量',TD024 AS '贈品量',TD025 AS '贈品已交量',(TD008-TD009+TD024-TD025) AS '總未出貨量',TD010 AS '品號單位',TD001 AS '訂單單別',TD002 AS '訂單單號',TD003 AS '訂單序號',TD016 AS '訂單狀態',MOCTA.TA001 AS '批次轉製令單別',MOCTA.TA002 AS '批次轉製令單號',MOCTA.TA009 AS '製令預計開工日',MOCTA.TA012 AS '製令實際開工日',MOCTA.TA010 AS '製令預計完工日' ,MOCTA.TA014 AS '製令實際完工日',MOCTA.TA015 AS '製令預計產量',MOCTA.TA017 AS '實際入庫數量'");
            SB.AppendFormat(" ,(CASE WHEN MOCTA.TA011='Y' THEN '已完工' ELSE CASE WHEN MOCTA.TA011='y' THEN '指定完工' ELSE  CASE WHEN MOCTA.TA011='1' THEN '未生產' ELSE CASE WHEN MOCTA.TA011='2' THEN '已發料' ELSE CASE WHEN MOCTA.TA011='3' THEN '生產中' ELSE '' END END END END END)AS '生產進度'");
            SB.AppendFormat(" ,(CASE WHEN CONVERT(datetime,MOCTA.TA009)<CONVERT(datetime,MOCTA.TA012) THEN '是' ELSE ''  END ) AS '製令開工異常警示'");
            SB.AppendFormat(" ,(CASE WHEN CONVERT(datetime,MOCTA.TA010)<CONVERT(datetime,MOCTA.TA014) THEN '是' ELSE ''  END ) AS '製令完工異常警示'");
            SB.AppendFormat(" ,(CASE WHEN MOCTA.TA017<MOCTA.TA015 THEN '是' ELSE ''  END) AS '產量不足'");
            SB.AppendFormat(" ,LRPTA.TA001 AS '批次計畫單號'");
            SB.AppendFormat(" ,(CASE WHEN ISNULL(MOCTA.TA033,'')<>''  THEN '是' ELSE ''  END )  AS '製令發放'");
            SB.AppendFormat(" ,(CASE WHEN CONVERT(datetime,TD013)<=CONVERT(datetime,MOCTA.TA009) THEN '是' ELSE ''  END )  AS '訂單是否延遲生產'");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" LEFT JOIN [TK].dbo.MOCTA ON MOCTA.TA026=TD001 AND MOCTA.TA027=TD002 AND MOCTA.TA028=TD003");
            SB.AppendFormat(" LEFT JOIN [TK].dbo.LRPTA ON LRPTA.TA023=TD001 AND LRPTA.TA024=TD002 AND LRPTA.TA025=TD003");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'",dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TD004 LIKE '4%'");
            SB.AppendFormat(" AND (TD008-TD009+TD024-TD025)>0");
            SB.AppendFormat(" AND TD021='Y' ");
            SB.AppendFormat(" AND TD016='N'");
            SB.AppendFormat(" AND TC001 IN ('A221', 'A222','A223','A228')");
            SB.AppendFormat(" ORDER BY TC001,TC053,TD013,TD004");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT5()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL5();
            Report report5 = new Report();
            report5.Load(@"REPORT\未出訂單業績統計.frx");

            report5.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report5.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report5.SetParameterValue("P1",  dateTimePicker7.Value.ToString("yyyyMMdd"));
            report5.SetParameterValue("P2",  dateTimePicker8.Value.ToString("yyyyMMdd"));
            report5.Preview = previewControl5;
            report5.Show();
        }

        public StringBuilder SETSQL5()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT '國內' AS '國別','劉莉琴' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'",dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TD016='N' AND TC006='140049' AND TC005='106000'");
            SB.AppendFormat(" GROUP BY TC008");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT '國內' AS '國別','蔡顏鴻' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}' ", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TD016='N' AND TC006='140078' AND TC005='106200'");
            SB.AppendFormat(" GROUP BY TC008");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT '大陸' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額'");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TD016='N' AND TC006='160155' AND TC005='106800'");
            SB.AppendFormat(" GROUP BY TC008");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT '國外' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額'");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TD016='N' AND TC006='160155' AND TC005='106300'");
            SB.AppendFormat("GROUP BY TC008 ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
            SETFASTREPORT2();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT5();
        }
        #endregion


    }
}
