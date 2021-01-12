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
using Calendar.NET;
using Excel = Microsoft.Office.Interop.Excel;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmREPORTMOCINVMB : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
      
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
       
        string tablename = null;
        int rownum = 0;

        SqlTransaction tran;

        DataSet ds1 = new DataSet();
        int result;

        Report report1 = new Report();

        public frmREPORTMOCINVMB()
        {
            InitializeComponent();

            SETDATES();
        }

        #region FUNCTION

        public void SETDATES()
        {
            DateTime FirstDay = DateTime.Now.AddDays(-DateTime.Now.Day + 1).AddMonths(-1);
            DateTime LastDay = DateTime.Now.AddMonths(1).AddDays(-DateTime.Now.AddMonths(1).Day).AddMonths(-1);

            dateTimePicker1.Value = FirstDay;
            dateTimePicker2.Value = LastDay;

        }
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1();

            Report report1 = new Report();
            report1.Load(@"REPORT\生產報表-得料率報表-品號.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" SELECT 
                                線別, SUBSTRING(製令單號, 1, 8) AS '日期', 品號, 品名, 規格, 製令單別, 製令單號, 生產單位, 預計產量, 生產量, 淨重, 單片重, 袋重, 袋重比, 蒸發率, 原料用量, 成品用量 / 1000 AS 成品用量, 類別, 領料是否扣袋重, 成品是否扣袋重
                                , CASE WHEN 領料是否扣袋重 IN('Y') AND 成品是否扣袋重 IN('Y') AND  類別 NOT IN('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) > 0 THEN(((生產量 * 淨重 * (1 - 袋重比))) / 1000) / (原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) ELSE 0 END  AS '領料扣成品扣的得料率'
                                , CASE WHEN 領料是否扣袋重 IN('Y') AND 成品是否扣袋重 IN('N') AND  類別 NOT IN('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) > 0 THEN(((生產量 * 淨重)) / 1000) / (原料用量 * (1 - 蒸發率) + (成品用量 / 1000) - (袋重比 * 原料用量)) ELSE 0 END AS '領料扣成品不扣的得料率'
                                , CASE WHEN 領料是否扣袋重 IN('N') AND 成品是否扣袋重 IN('Y') AND  類別 NOT IN('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) > 0 THEN(((生產量 * 淨重 * (1 - 袋重比))) / 1000) / (原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) ELSE 0 END AS '領料不扣成品扣的得料率'
                                , CASE WHEN 領料是否扣袋重 IN('N') AND 成品是否扣袋重 IN('N') AND  類別 NOT IN('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) > 0 THEN(((生產量 * 淨重) / 1000) / (原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000))) ELSE 0 END AS '領料不扣成品不扣的得料率'
                                , CASE WHEN 類別 IN('半成品') AND 原料用量 > 0  AND 成品是否扣袋重 IN('Y') THEN(生產量 - (生產量 * 袋重比)) / (原料用量 * (1 - 蒸發率 / 100)) ELSE 0 END  AS '半成品得料率(成品扣袋重)'
                                , CASE WHEN 類別 IN('半成品') AND 原料用量 > 0  AND 成品是否扣袋重 IN('N') THEN(生產量) / (原料用量 * (1 - 蒸發率 / 100)) ELSE 0 END  AS '半成品得料率(成品不扣袋重)'
                                , CASE WHEN 類別 IN('個', '試吃') AND 原料用量 > 0 AND(原料用量 * (1 - (蒸發率 / 100))) > 0 THEN(生產量 * 淨重 / 1000) / (原料用量 * (1 - (蒸發率 / 100))) ELSE 0 END  AS '個/試吃得料率'
                                , CASE WHEN 類別 IN('片') AND 原料用量 > 0 AND(原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) > 0 THEN(生產量 * 淨重 / 1000) / (原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) ELSE 0 END  AS '片得料率'
                                , CASE WHEN 類別 IN('單包') AND 原料用量 > 0 THEN 生產量 / 原料用量  ELSE 0 END AS '單包得料率'
                                , CASE WHEN 類別 IN('kg') AND(原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) > 0 THEN((生產量) / (原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (原料用量 * 袋重比))) ELSE 0 END AS 'kg得料率'
                                , (CASE WHEN 領料是否扣袋重 IN('Y') AND 成品是否扣袋重 IN('Y') AND  類別 NOT IN('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) > 0 THEN(((生產量 * 淨重 * (1 - 袋重比))) / 1000) / (原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) ELSE 0 END) + (CASE WHEN 領料是否扣袋重 IN('Y') AND 成品是否扣袋重 IN('N') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) > 0 THEN(((生產量 * 淨重)) / 1000) / (原料用量 * (1 - 蒸發率) + (成品用量 / 1000) - (袋重比 * 原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN('N') AND 成品是否扣袋重 IN('Y') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) > 0 THEN(((生產量 * 淨重 * (1 - 袋重比))) / 1000) / (原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN('N') AND 成品是否扣袋重 IN('N') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) > 0 THEN(((生產量 * 淨重) / 1000) / (原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000))) ELSE 0 END)+(CASE WHEN 類別 IN('半成品') AND 原料用量> 0  AND 成品是否扣袋重 IN('Y') THEN(生產量 - (生產量 * 袋重比)) / (原料用量 * (1 - 蒸發率 / 100)) ELSE 0 END)+(CASE WHEN 類別 IN('半成品') AND 原料用量> 0  AND 成品是否扣袋重 IN('N') THEN(生產量) / (原料用量 * (1 - 蒸發率 / 100)) ELSE 0 END)+(CASE WHEN 類別 IN('個', '試吃') AND 原料用量> 0 AND(原料用量 * (1 - (蒸發率 / 100))) > 0 THEN(生產量 * 淨重 / 1000) / (原料用量 * (1 - (蒸發率 / 100))) ELSE 0 END)+(CASE WHEN 類別 IN('片') AND 原料用量> 0 AND(原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) > 0 THEN(生產量 * 淨重 / 1000) / (原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) ELSE 0 END)+(CASE WHEN 類別 IN('單包') AND 原料用量> 0 THEN 生產量/ 原料用量  ELSE 0 END)+(CASE WHEN 類別 IN('kg') AND(原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) > 0 THEN((生產量) / (原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (原料用量 * 袋重比))) ELSE 0 END) AS '得料率'
                                ,(CASE WHEN 領料是否扣袋重 IN('Y') AND 成品是否扣袋重 IN('Y') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) > 0 THEN(((生產量 * 淨重 * (1 - 袋重比))) / 1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN('Y') AND 成品是否扣袋重 IN('N') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) > 0 THEN(((生產量 * 淨重)) / 1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN('N') AND 成品是否扣袋重 IN('Y') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) > 0 THEN(((生產量 * 淨重 * (1 - 袋重比))) / 1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN('N') AND 成品是否扣袋重 IN('N') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) > 0 THEN((生產量 * 淨重) / 1000) ELSE 0 END)+(CASE WHEN 類別 IN('半成品') AND 原料用量> 0  AND 成品是否扣袋重 IN('Y') THEN(生產量 - (生產量 * 袋重比)) ELSE 0 END )+(CASE WHEN 類別 IN('半成品') AND 原料用量> 0  AND 成品是否扣袋重 IN('N') THEN(生產量) ELSE 0 END)+(CASE WHEN 類別 IN('個', '試吃') AND 原料用量> 0 AND(原料用量 * (1 - (蒸發率 / 100))) > 0 THEN(生產量 * 淨重 / 1000) ELSE 0 END)+(CASE WHEN 類別 IN('片') AND 原料用量> 0 AND(原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) > 0 THEN(生產量 * 淨重 / 1000) ELSE 0 END)+(CASE WHEN 類別 IN('單包') AND 原料用量> 0 THEN 生產量  ELSE 0 END)+(CASE WHEN 類別 IN('kg') AND(原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) > 0 THEN((生產量)) ELSE 0 END) AS '分子'
                                ,(CASE WHEN 領料是否扣袋重 IN('Y') AND 成品是否扣袋重 IN('Y') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) > 0 THEN(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN('Y') AND 成品是否扣袋重 IN('N') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) > 0 THEN(原料用量 * (1 - 蒸發率) + (成品用量 / 1000) - (袋重比 * 原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN('N') AND 成品是否扣袋重 IN('Y') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (袋重比 * 原料用量)) > 0 THEN(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN('N') AND 成品是否扣袋重 IN('N') AND 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') AND(原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000)) > 0 THEN((原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000))) ELSE 0 END)+(CASE WHEN 類別 IN('半成品') AND 原料用量> 0  AND 成品是否扣袋重 IN('Y') THEN(原料用量 * (1 - 蒸發率 / 100)) ELSE 0 END)+(CASE WHEN 類別 IN('半成品') AND 原料用量> 0  AND 成品是否扣袋重 IN('N') THEN(原料用量 * (1 - 蒸發率 / 100)) ELSE 0 END)+(CASE WHEN 類別 IN('個', '試吃') AND 原料用量> 0 AND(原料用量 * (1 - (蒸發率 / 100))) > 0 THEN(原料用量 * (1 - (蒸發率 / 100))) ELSE 0 END)+(CASE WHEN 類別 IN('片') AND 原料用量> 0 AND(原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) > 0 THEN(原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) ELSE 0 END)+(CASE WHEN 類別 IN('單包') AND 原料用量> 0 THEN 原料用量  ELSE 0 END)+(CASE WHEN 類別 IN('kg') AND(原料用量 * (1 - (蒸發率 / 100)) - (原料用量 * 袋重比)) > 0 THEN((原料用量 * (1 - (蒸發率 / 100)) + (成品用量 / 1000) - (原料用量 * 袋重比))) ELSE 0 END) AS '分母'


                                FROM(
                                SELECT MD002 AS '線別', TA006 AS '品號', TA034 AS '品名', MB003 AS '規格', TA001 AS '製令單別', TA002 AS '製令單號', TA007 AS '生產單位', INVMB.UDF03 AS '類別', TA015 AS '預計產量', TA017 AS '生產量', INVMB.UDF07 AS '淨重', INVMB.UDF08 AS '單片重', INVMB.UDF09 AS '袋重', INVMB.UDF06 AS '蒸發率', INVMB.UDF01 AS '成品是否扣袋重', INVMB.UDF02 AS '領料是否扣袋重'
                                , (SELECT ISNULL(SUM(TB005), 0) FROM[TK].dbo.MOCTB TB WHERE(TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TB.TB001 = MOCTA.TA001 AND TB.TB002 = MOCTA.TA002)  AS '原料用量'
                                ,(SELECT ISNULL(SUM(TB005 * MB.UDF07), 0) FROM[TK].dbo.MOCTB TB,[TK].dbo.INVMB MB WHERE TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TB.TB001= MOCTA.TA001 AND TB.TB002= MOCTA.TA002) AS '成品用量'
                                ,CASE WHEN INVMB.UDF08>0 AND INVMB.UDF09>0  THEN 1/(INVMB.UDF08+INVMB.UDF09)*INVMB.UDF09 ELSE 0 END AS '袋重比'
                                FROM[TK].dbo.INVMB,[TK].dbo.MOCTA,[TK].dbo.CMSMD
                                WHERE TA006=MB001 AND TA021=MD001
                                AND ISNULL(INVMB.UDF03,'')<>''
                                AND TA003>='{0}' AND TA003<='{1}'
                                AND TA006 = '{2}'


                                UNION ALL
                                SELECT MD002 AS '線別', TA006 AS '品號', TA034 AS '品名', MB003 AS '規格', TA001 AS '製令單別', TA002 AS '製令單號', TA007 AS '生產單位', INVMB.UDF03 AS '類別', TA015 AS '預計產量', TA017 AS '生產量', INVMB.UDF07 AS '淨重', INVMB.UDF08 AS '單片重', INVMB.UDF09 AS '袋重', INVMB.UDF06 AS '蒸發率', INVMB.UDF01 AS '成品是否扣袋重', INVMB.UDF02 AS '領料是否扣袋重'
                                , (SELECT ISNULL(SUM(TB005), 0) FROM[TK].dbo.MOCTB TB WHERE(TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TB.TB001 = MOCTA.TA001 AND TB.TB002 = MOCTA.TA002)  AS '原料用量'
                                ,(SELECT ISNULL(SUM(TB005 * MB.UDF07), 0) FROM[TK].dbo.MOCTB TB,[TK].dbo.INVMB MB WHERE TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TB.TB001= MOCTA.TA001 AND TB.TB002= MOCTA.TA002) AS '成品用量'
                                ,CASE WHEN INVMB.UDF08>0 AND INVMB.UDF09>0  THEN 1/(INVMB.UDF08+INVMB.UDF09)*INVMB.UDF09 ELSE 0 END AS '袋重比'
                                FROM[TK].dbo.INVMB,[TK].dbo.MOCTA,[TK].dbo.CMSMD
                                WHERE TA006=MB001 AND TA021=MD001
                                AND ISNULL(INVMB.UDF03,'')<>''
                                AND TA003>='{0}' AND TA003<='{1}'
                                AND TA006 IN (SELECT MD003 FROM [TK].dbo.BOMMD WHERE MD003 LIKE '3%' AND  MD001='{2}')

                                ) AS TEMP
                                ORDER BY 線別,SUBSTRING(製令單號,1,8),品號", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), textBox1.Text.Trim());


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
