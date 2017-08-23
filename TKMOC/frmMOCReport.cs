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

namespace TKMOC
{
    public partial class frmMOCReport : Form
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

        public frmMOCReport()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    //dataGridView1.Columns.Clear();
                    ds.Clear();
                    
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();

                    if (ds.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();                        
                        //dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];
                        
                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }

        }

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();
            DateTime dt = dateTimePicker1.Value;
            string ThisYear = null;
            string ThisMonth = null;
            string LastMonth = null;
            string LastYear = null;
            string LastYearMonth = null;

            ThisYear = dateTimePicker1.Value.ToString("yyyy");
            ThisMonth = dateTimePicker1.Value.ToString("MM");
            LastMonth = dt.AddMonths(-1).ToString("MM");
            LastYear = dt.AddYears(-1).ToString("yyyy");
            LastYearMonth = dt.AddYears(-1).AddMonths(1).ToString("MM");

            if (comboBox1.Text.ToString().Equals("生產日報的分析表"))
            {
                
                STR.AppendFormat(@"  SELECT ");
                STR.AppendFormat(@"  [PRODUCEDATE] AS '日期',[PRODUCEDEP] AS '線別',[PRODUCENAME] AS '品名'");
                STR.AppendFormat(@"  ,[WEIGHTBEFORECOOK] AS '總投入量',[REWORKPCT] AS '重工佔比',[EVARATE] AS '蒸發率'");
                STR.AppendFormat(@"  ,[STIRPCT] AS '攪拌成型率',[MANULOST]	 AS '製成損失率',[PCT] AS '餅製成率'");
                STR.AppendFormat(@"  ,[TOTALPCT] AS '總製成率',[CANPCT] AS '罐裝製成率',[STIR] AS '攪拌不良'");
                STR.AppendFormat(@"  ,[SIDES]	 AS '成型邊料',[COOKIES] AS '餅麩',[COOK] AS '烤焙',[NGPACKAGE] AS '包裝不良餅乾'");
                STR.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]");
                STR.AppendFormat(@"  WHERE  [PRODUCEDATE]>='{0}' AND [PRODUCEDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  UNION ALL");
                STR.AppendFormat(@"  SELECT ");
                STR.AppendFormat(@"  '9999/9/9','小計','小計'");
                STR.AppendFormat(@"  ,SUM([WEIGHTBEFORECOOK]) AS '總投入量',AVG([REWORKPCT]) AS '重工佔比'");
                STR.AppendFormat(@"  ,AVG([EVARATE]) AS '蒸發率',AVG([STIRPCT]) AS '攪拌成型率'");
                STR.AppendFormat(@"  ,AVG([MANULOST])	 AS '製成損失率',AVG([PCT]) AS '餅製成率'");
                STR.AppendFormat(@"  ,AVG([TOTALPCT]) AS '總製成率',AVG([CANPCT]) AS '罐裝製成率'");
                STR.AppendFormat(@"  ,SUM([STIR]) AS '攪拌不良',SUM([SIDES])	 AS '成型邊料',SUM([COOKIES]) AS '餅麩'");
                STR.AppendFormat(@"  ,SUM([COOK]) AS '烤焙',SUM([NGPACKAGE]) AS '包裝不良餅乾'");
                STR.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]");
                STR.AppendFormat(@"  WHERE  [PRODUCEDATE]>='{0}' AND [PRODUCEDATE]<='{1}'",dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY [PRODUCEDATE],[PRODUCEDEP],[PRODUCENAME]");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
               

                tablename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals("生產日報的月份分析表"))
            {
                string YEARS = dateTimePicker1.Value.ToString("yyyy");
                STR.AppendFormat(@"  SELECT '{0}' AS '年度', [ID]  AS '月份'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT SUM([WEIGHTBEFORECOOK]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])='{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '總投入量'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT AVG([REWORKPCT]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '重工佔比'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT AVG([EVARATE]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '蒸發率'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT AVG([STIRPCT]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '攪拌成型率'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT AVG([MANULOST]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '製成損失率'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT AVG([PCT]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '餅製成率'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT AVG([TOTALPCT]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])='{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '總製成率'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT AVG([CANPCT])  FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '罐裝製成率'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT SUM([STIR])  FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '攪拌不良'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT SUM([SIDES]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '成型邊料'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT SUM([COOKIES]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '餅麩'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT SUM([COOK]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '烤焙'", YEARS);
                STR.AppendFormat(@"  ,ISNULL((SELECT SUM([NGPACKAGE]) FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT] WHERE YEAR([PRODUCEDATE])= '{0}' AND MONTH([PRODUCEDATE])=[BASEMONTH].[ID]),0) AS '包裝不良餅乾'", YEARS);
                STR.AppendFormat(@"  FROM [TKMOC].[dbo].[BASEMONTH]");
                STR.AppendFormat(@"  ");



                tablename = "TEMPds2";
            }
            else if (comboBox1.Text.ToString().Equals("不良品餅乾報廢明細表"))
            {
                STR.AppendFormat(@"  SELECT [MAIN] AS '線別',CONVERT(varchar(100),[MAINDATE], 111) AS '日期',[DAMAGEDCOOKIES] AS '破損餅乾(kg)',[LANDCOOKIES] AS '落地餅乾(kg)',[SCRAPCOOKIES]  AS '餅乾屑(kg)',[ID]");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSCRAPPEDMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'",dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY LEN([MAIN]),[MAIN],CONVERT(varchar(100),[MAINDATE], 111)");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds3";
            }
            else if (comboBox1.Text.ToString().Equals("生產日報表明細表"))
            {
                STR.AppendFormat(@"  SELECT ");
                STR.AppendFormat(@"  [PRODUCETYPE] AS '成品/半成品' ,[PRODUCEDEP] AS '製造組'");
                STR.AppendFormat(@"  ,[PRODUCEDATE] AS '日期',[PRODUCEMB001] AS '品號',[PRODUCENAME] AS '品名' ");
                STR.AppendFormat(@"  ,[PASTRYPREIN] AS '油酥預計投入量(kg)',[PASTRY] AS '油酥原料'");
                STR.AppendFormat(@"  ,[PASTRYRECYCLE] AS '油酥可回收餅麩' ,[WATERFLOURPREIN] AS '水麵預計投入量(kg)'");
                STR.AppendFormat(@"  ,[WATERFLOUR] AS '水面原料',[WATERFLOURSIDE] AS '水面可回收邊料' ");
                STR.AppendFormat(@"  ,[WATERFLOURRECYCLE] AS '水面可回收餅麩',[PASTRYFLODTIME] AS '油酥、摺疊製造時間(分)'");
                STR.AppendFormat(@"  ,[PASTRYFLODNUM] AS '油酥、摺疊製造人數' ,[WATERFLOURTIME] AS '水面製造時間(分)'");
                STR.AppendFormat(@"  ,[WATERFLOURNUM] AS '水面製造人數',[RECYCLEFLOUR] AS '今日產生可回收餅麩'");
                STR.AppendFormat(@"  ,[KNIFENUM] AS '刀數',[WEIGHTBEFRORE] AS '烤前單片重量(g)'");
                STR.AppendFormat(@"  ,[WEIGHTAFTER] AS '烤後單片重量(g)' ,[ROWNUM] AS '每排數量'");
                STR.AppendFormat(@"  ,[NGTOTAL] AS '未熟總量(kg)'");
                STR.AppendFormat(@"  ,[NGCOOKTIME] AS '未熟烤焙時間(分)' ,[RECOOKTIME] AS '重烤重工時間',[PREOUT] AS '預計產出(kg)'");
                STR.AppendFormat(@"  ,[PACKAGETIME] AS '包裝時間(內包裝區/罐裝)(分)',[PACKAGENUM] AS '包裝人數' ");
                STR.AppendFormat(@"  ,[STIR] AS '攪拌',[SIDES] AS '成型邊料(kg)',[COOKIES] AS '餅麩(kg)'");
                STR.AppendFormat(@"  ,[COOK] AS '篩選餅乾區不良烤焙(kg)'");
                STR.AppendFormat(@" ,[OUTCOOKIES] AS '篩選餅乾區餅乾屑(kg)' ,[CLEANCOOKIES] AS '清掃廢料(kg)'  ");
                STR.AppendFormat(@"  ,[NGPACKAGE] AS '包裝不良餅乾(kg)'");
                STR.AppendFormat(@"  ,[NGPACKAGECAN] AS '包裝(內袋(卷) 罐)',[CAN] AS '包裝投入(袋(卷),罐)'");
                STR.AppendFormat(@"  ,[WEIGHTCAN] AS '一箱裸餅重' ,[WEIGHTCANBOXED] AS '一箱餅含袋重'");
                STR.AppendFormat(@"  ,[HLAFWEIGHT] AS '半成品入庫數(kg) (含袋重)',[REMARK] AS '備註' ");
                STR.AppendFormat(@"  ,[MANUTIME] AS '製造工時(分)',[PACKTIME] AS '包裝工時(分)'");
                STR.AppendFormat(@"  ,[WEIGHTBEFORECOOK] AS '烤前實際總投入 (kg)'  ,[WEIGHTAFTERCOOK] AS '烤後實際總投入 (kg)'");
                STR.AppendFormat(@"  ,[ACTUALOUT] AS '實際產出(kg)(裸餅)',[WEIGHTPACKAGE] AS '袋重(kg)' ");
                STR.AppendFormat(@"  ,[PACKLOST] AS '包裝損耗率',[HLAFLOST] AS '半成品產出效率'");
                STR.AppendFormat(@"  ,[REWORKPCT] AS '重工佔比',[TOTALTIME] AS '總工時(分)' ");
                STR.AppendFormat(@"  ,[STIRPCT] AS '攪拌成型製成率%',[EVARATE] AS '蒸發率'");
                STR.AppendFormat(@"  ,[MANULOST] AS '製成損失率',[PCT] AS '製成率' ,[PRETIME] AS '前置時間'");
                STR.AppendFormat(@"  ,[STOPTIME] AS '停機時間' ,[PREWEIGT] AS '容量規格'");
                STR.AppendFormat(@"  ,[PRECAN] AS '預計包罐數',[ACTUALCAN] AS '實際包罐數',[TOTALPCT] AS '總製成率'");
                STR.AppendFormat(@"  ,[CANPCT] AS '總包罐製成率',TRYCAN AS '預計試吃包罐數'");
                STR.AppendFormat(@"  ,ACTUALTRYCAN  AS '實際試吃包罐數'  ,[ID]  ");
                STR.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCPRODUCTDAILYREPORT]   WITH (NOLOCK) ");
                STR.AppendFormat(@"  WHERE [PRODUCEDATE] >='{0}' AND [PRODUCEDATE] <='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY [PRODUCEDATE],[ID]  ");
                STR.AppendFormat(@"  ");
       


                tablename = "TEMPds4";
            }
            else if (comboBox1.Text.ToString().Equals("不良餅麩明細表"))
            {
                STR.AppendFormat(@"    SELECT 日期,時間,品名,回收量,不良品報廢,線別,品號,單別,單號");
                STR.AppendFormat(@"    FROM (");
                STR.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[MAINDATE],112) AS '日期',CONVERT(varchar(100),[MAINTIME],8)  AS '時間',[MB002] AS '品名',[NUM] AS '回收量',[NGNUM] AS '不良品報廢' ,[MAIN] AS '線別',[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGCOOKIESMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  UNION ALL ");
                STR.AppendFormat(@"  SELECT  '9999/9/9','合計','',SUM([NUM]) AS '回收量',SUM([NGNUM]) AS '不良品報廢' , '新廠製全部組','','',''");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGCOOKIESMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ) AS TEMP");
                STR.AppendFormat(@"    ORDER BY LEN(線別),線別,日期,品號 ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds5";
            }
            else if (comboBox1.Text.ToString().Equals("不良邊料明細表"))
            {
                STR.AppendFormat(@"   SELECT 日期,時間,品名,回收邊料,不良品報廢,線別,品號,單別,單號");
                STR.AppendFormat(@"  FROM (");
                STR.AppendFormat(@"  SELECT  CONVERT(NVARCHAR,[MAINDATE],112) AS '日期', CONVERT(varchar(100),[MAINTIME],8)  AS '時間',[MB002] AS '品名',[NUM] AS '回收邊料',[NGNUM] AS '不良品報廢',[MAIN] AS '線別',[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSIDEMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  UNION ALL ");
                STR.AppendFormat(@"  SELECT  '9999/9/9','合計','',SUM([NUM]) AS '回收邊料',SUM([NGNUM]) AS '不良品報廢' , '新廠製造全組****','','',''");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSIDEMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ) AS TEMP ");
                STR.AppendFormat(@"  ORDER BY LEN(線別),線別,日期,品號");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds6";
            }
            else if (comboBox1.Text.ToString().Equals("不良未熟明細表"))
            {

                STR.AppendFormat(@"   SELECT 日期,時間,品名,未熟餅,烤培時間,不良品報廢,線別,品號,單別,單號");
                STR.AppendFormat(@"   FROM (");
                STR.AppendFormat(@"  SELECT  [MAINDATE] AS '日期',CONVERT(varchar(100),[MAINTIME],8) AS '時間',[MB002] AS '品名',[NUM] AS '未熟餅',[COOKTIME] AS '烤培時間',[NGNUM] AS '不良品報廢',[MAIN] AS '線別',[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGNOBURNMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  UNION ALL ");
                STR.AppendFormat(@"  SELECT  '9999/9/9','合計','',SUM([NUM]) AS '未熟餅',SUM([COOKTIME]) AS '烤培時間',SUM([NGNUM]) AS '不良品報廢3' , '新廠製造全組****','','',''");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGNOBURNMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ) AS TEMP ");
                STR.AppendFormat(@"  ORDER BY LEN(線別),線別,日期,品號 ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds7";
            }
            else if (comboBox1.Text.ToString().Equals("烤爐溫度明細表"))
            {
                STR.AppendFormat(@"  SELECT CONVERT(varchar(100),[OVENDATE], 112) AS '日期',[MANUDEP].[DEPNAME] AS '組',CONVERT(varchar(100),[PREHEARTSTART], 108)  AS '預熱時間(起)',CONVERT(varchar(100),[PREHEARTEND], 108)   AS '預熱時間(迄)',[GAS]  AS '瓦斯磅數',EMP1.NAME  AS '折疊人員1',EMP2.NAME    AS '折疊人員2', EMP3.NAME   AS '主管',EMP4.NAME    AS '操作人員',[MANUDEP] AS '組別'");
                STR.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[TEMPER] AS '溫度',[HUMIDITY] AS '溼度',[WEATHER] AS '天氣',CONVERT(varchar(100),[MANUTIME],108) AS '時間'");
                STR.AppendFormat(@" ,[FURANACEUP1] AS '上爐1-1',[FURANACEUP2] AS '上爐2-1',[FURANACEUP3] AS '上爐3-1',[FURANACEUP4] AS '上爐4-1',[FURANACEUP5] AS '上爐5-1'");
                STR.AppendFormat(@" ,[FURANACEUP1A] AS '上爐1-2',[FURANACEUP2A] AS '上爐2-2',[FURANACEUP3A] AS '上爐3-2',[FURANACEUP4A] AS '上爐4-2',[FURANACEUP5A] AS '上爐5-2'");
                STR.AppendFormat(@" ,[FURANACEUP1B] AS '上爐1-3',[FURANACEUP2B] AS '上爐2-3',[FURANACEUP3B] AS '上爐3-3',[FURANACEUP4B] AS '上爐4-3',[FURANACEUP5B] AS '上爐5-3' ");
                STR.AppendFormat(@" ,[FURANACEDOWN1] AS '下爐1-1',[FURANACEDOWN2] AS '下爐2-1',[FURANACEDOWN3] AS '下爐3-1',[FURANACEDOWN4] AS '下爐4-1',[FURANACEDOWN5] AS '下爐5-1'");
                STR.AppendFormat(@" ,[FURANACEDOWN1A] AS '下爐1-2',[FURANACEDOWN2A] AS '下爐2-2',[FURANACEDOWN3A] AS '下爐3-2',[FURANACEDOWN4A] AS '下爐4-2',[FURANACEDOWN5A] AS '下爐5-2'");
                STR.AppendFormat(@" ,[FURANACEDOWN1B] AS '下爐1-3',[FURANACEDOWN2B] AS '下爐2-3',[FURANACEDOWN3B] AS '下爐3-3',[FURANACEDOWN4B] AS '下爐4-3',[FURANACEDOWN5B] AS '下爐5-3'");
                STR.AppendFormat(@"  ,[MOCOVENDTAIL].[ID],[SOURCEID]");
                STR.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCOVENDTAIL], [TKMOC].[dbo].[MOCOVEN]");
                STR.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MANUEMPLOYEE] EMP1  ON [FLODPEOPLE1]=EMP1.ID");
                STR.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MANUEMPLOYEE] EMP2 ON [FLODPEOPLE2]=EMP2.ID");
                STR.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MANUEMPLOYEE]  EMP3 ON [MANAGER]=EMP3.ID");
                STR.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MANUEMPLOYEE]  EMP4 ON [OPERATOR]=EMP4.ID");
                STR.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MANUDEP] ON [MANUDEP].ID=[MOCOVEN].[MANUDEP]");
                STR.AppendFormat(@"  WHERE [MOCOVENDTAIL].[SOURCEID]=[MOCOVEN].[ID]");
                STR.AppendFormat(@"  AND CONVERT(varchar(100),[OVENDATE], 112)>='{0}' AND CONVERT(varchar(100),[OVENDATE], 112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY CONVERT(varchar(100),[OVENDATE], 112),[MANUDEP].[DEPNAME],[MB001]");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds8";
            }
            else if (comboBox1.Text.ToString().Equals("成型檢驗表"))
            {
                STR.AppendFormat(@"  SELECT [CHECKCOOKIESMD].[MB002] AS '品名',CONVERT(varchar(100),[STIME],8) AS '開始時間',CONVERT(varchar(100),[ETIME],8) AS '結束時間',[SLOT] AS '桶數',[CUTNUMBER] AS '刀數',[WEIGHT] AS '重量',[CHECKCOOKIESMD].[MAIN] AS '線別',[CHECKCOOKIESMD].[MAINDATE] AS '日期'");
                STR.AppendFormat(@"  ,[CHECKCOOKIESMD].[TARGETPROTA001] AS '單別',[CHECKCOOKIESMD].[TARGETPROTA002] AS '單號',[CHECKCOOKIESMD].[MB001] AS '品號'");
                STR.AppendFormat(@"  ,CONVERT(varchar(100),[CHECKTIME],8) AS '時間',[WIGHT] AS '重量',[LENGTH] AS '長度',[TEMP] AS '溫度',[HUMIDITY] AS '溼度',[CHECKRESULT] AS '檢查結果',[OWNER] AS '填表人',[MANAGER]  AS '主管'");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKCOOKIESM],[TKCIM].[dbo].[CHECKCOOKIESMD]");
                STR.AppendFormat(@"  WHERE [CHECKCOOKIESM].[TARGETPROTA001]=[CHECKCOOKIESMD].[TARGETPROTA001] AND [CHECKCOOKIESM].[TARGETPROTA002]=[CHECKCOOKIESMD].[TARGETPROTA002] ");
                STR.AppendFormat(@"  AND CONVERT(varchar(100),[CHECKCOOKIESMD].[MAINDATE], 112)>='{0}' AND CONVERT(varchar(100),[CHECKCOOKIESMD].[MAINDATE], 112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY [CHECKCOOKIESMD].[MAINDATE],[CHECKCOOKIESMD].[MAIN],CONVERT(varchar(100),[CHECKTIME],8)");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds9";
            }
            else if (comboBox1.Text.ToString().Equals("水麵添加表"))
            {
                STR.AppendFormat(@"  SELECT [MAIN] AS '組別',CONVERT(NVARCHAR,[MAINDATE],112)  AS '生產日'  ,[MATERWATERPROIDM].[TARGETPROTA001] AS '單別'");
                STR.AppendFormat(@"  ,[MATERWATERPROIDM].[TARGETPROTA002] AS '單號'  ,[MATERWATERPROIDM].[MB001] AS '品號'");
                STR.AppendFormat(@"  ,[MATERWATERPROIDM].[MB002] AS '品名',[MATERWATERPROIDM].[LOTID] AS '批號'  ,[CANNO] AS '桶數'");
                STR.AppendFormat(@"  ,[NUM] AS '重量'  ,[OUTLOOK] AS '外觀',CONVERT(varchar(100),[STIME],8) AS '起時間'");
                STR.AppendFormat(@"  ,CONVERT(varchar(100),[ETIME],8) AS '迄時間'  ,[TEMP] AS '溫度' ,[HUDI] AS '溼度',[MOVEIN] AS '投料人'");
                STR.AppendFormat(@"  ,[CHECKEMP] AS '抽檢人'  ");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPROIDM]");
                STR.AppendFormat(@"  LEFT JOIN [TKCIM].[dbo].[MATERWATERPROIDMD]  ON [MATERWATERPROIDM].[TARGETPROTA001]=[MATERWATERPROIDMD].[TARGETPROTA001]   AND [MATERWATERPROIDM].[TARGETPROTA002]=[MATERWATERPROIDMD].[TARGETPROTA002]  AND [MATERWATERPROIDM].[MB001]=[MATERWATERPROIDMD].[MB001]   AND [MATERWATERPROIDM].[LOTID]=[MATERWATERPROIDMD].[LOTID]  ");
                STR.AppendFormat(@"  WHERE [MAINDATE]>= '{0}' AND [MAINDATE]<= '{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY LEN([MATERWATERPROIDM].[MAIN]),[MATERWATERPROIDM].[MAIN],[MATERWATERPROIDM].[TARGETPROTA001] ,[MATERWATERPROIDM].[TARGETPROTA002],CONVERT(INT,[CANNO]),[MATERWATERPROIDM].[MB001],[MATERWATERPROIDM].[LOTID]  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds10";
            }
            else if (comboBox1.Text.ToString().Equals("油酥添加表"))
            {
                STR.AppendFormat(@"  SELECT [MAIN] AS '組別',CONVERT(NVARCHAR,[MAINDATE],112)  AS '生產日'  ,[METEROILPROIDM].[TARGETPROTA001] AS '單別'");
                STR.AppendFormat(@"  ,[METEROILPROIDM].[TARGETPROTA002] AS '單號'  ,[METEROILPROIDM].[MB001] AS '品號'");
                STR.AppendFormat(@"  ,[METEROILPROIDM].[MB002] AS '品名',[METEROILPROIDM].[LOTID] AS '批號'  ,[CANNO] AS '桶數'");
                STR.AppendFormat(@"  ,[NUM] AS '重量'  ,[OUTLOOK] AS '外觀',CONVERT(varchar(100),[STIME],8) AS '起時間'");
                STR.AppendFormat(@"  ,CONVERT(varchar(100),[ETIME],8) AS '迄時間'  ,[TEMP] AS '溫度' ,[HUDI] AS '溼度'");
                STR.AppendFormat(@"  ,[MOVEIN] AS '投料人',[CHECKEMP] AS '抽檢人' ");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILPROIDM]");
                STR.AppendFormat(@"  LEFT JOIN [TKCIM].[dbo].[METEROILPROIDMD]  ON [METEROILPROIDM].[TARGETPROTA001]=[METEROILPROIDMD].[TARGETPROTA001]    AND [METEROILPROIDM].[TARGETPROTA002]=[METEROILPROIDMD].[TARGETPROTA002]    AND [METEROILPROIDM].[MB001]=[METEROILPROIDMD].[MB001]    AND [METEROILPROIDM].[LOTID]=[METEROILPROIDMD].[LOTID]  ");
                STR.AppendFormat(@"  WHERE [MAINDATE]>= '{0}' AND [MAINDATE]<= '{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY LEN([METEROILPROIDM].[MAIN]),[METEROILPROIDM].[MAIN],[METEROILPROIDM].[MAINDATE],[METEROILPROIDM].[TARGETPROTA001],[METEROILPROIDM].[TARGETPROTA002], CONVERT(INT,[CANNO])");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds11";
            }
            else if (comboBox1.Text.ToString().Equals("成型檢驗表"))
            {

                STR.AppendFormat(@"  SELECT [CHECKCOOKIESMD].[MB002] AS '品名',CONVERT(varchar(100),[CHECKTIME],8) AS '時間',[WIGHT] AS '重量'");
                STR.AppendFormat(@"  ,[CHECKCOOKIESMD].[LENGTH] AS '長度',[CHECKCOOKIESMD].[TEMP] AS '溫度',[CHECKCOOKIESMD].[HUMIDITY] AS '溼度',[CHECKCOOKIESMD].[CHECKRESULT] AS '檢查結果'");
                STR.AppendFormat(@"  ,[CHECKCOOKIESMD].[OWNER] AS '填表人',[CHECKCOOKIESMD].[MANAGER]  AS '主管',[CHECKCOOKIESMD].[MAIN] AS '線別',CONVERT(NVARCHAR,[CHECKCOOKIESMD].[MAINDATE],112) AS '日期'");
                STR.AppendFormat(@"  ,[CHECKCOOKIESMD].[TARGETPROTA001] AS '單別',[CHECKCOOKIESMD].[TARGETPROTA002] AS '單號',[CHECKCOOKIESMD].[MB001] AS '品號'");
                STR.AppendFormat(@"  ,[CHECKCOOKIESMD].[ID] ");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKCOOKIESM],[TKCIM].[dbo].[CHECKCOOKIESMD]");
                STR.AppendFormat(@"  WHERE [CHECKCOOKIESM].TARGETPROTA001=[CHECKCOOKIESMD].TARGETPROTA001 AND [CHECKCOOKIESM].TARGETPROTA002=[CHECKCOOKIESMD].TARGETPROTA002");
                STR.AppendFormat(@"  AND CONVERT(NVARCHAR,[CHECKCOOKIESMD].[MAINDATE],112)>='{0}' AND CONVERT(NVARCHAR,[CHECKCOOKIESMD].[MAINDATE],112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY [CHECKCOOKIESMD].[MAIN],CONVERT(NVARCHAR,[CHECKCOOKIESMD].[MAINDATE],112),[CHECKCOOKIESMD].[MB002]");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds12";
            }
            else if (comboBox1.Text.ToString().Equals("出爐餅溫量測記錄表"))
            {

                STR.AppendFormat(@"  SELECT [MB002] AS '品名',CONVERT(NVARCHAR,[CHECKTIME],8) AS '時間',[TEMP] AS '溫度',[OWNER] AS '檢測員',[MANAGER] AS '主管',[MAIN] AS '線別',[MAINDATE] AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號',[ID] ");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKBAKEDTEMPM]");
                STR.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[MAINDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MAINDATE],112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY [MAIN],CONVERT(NVARCHAR,[MAINDATE],112),CONVERT(NVARCHAR,[CHECKTIME],8)");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds13";
            }
            else if (comboBox1.Text.ToString().Equals("烘烤製程記錄"))
            {

                STR.AppendFormat(@"  SELECT ");
                STR.AppendFormat(@"  [CHECKOVENMD].[MB002] AS '品名',[TEMPER] AS '溫度',[HUMIDITY] AS '溼度',[WEATHER] AS '天氣',CONVERT(varchar(100),[MANUTIME], 8)  AS '時間'");
                STR.AppendFormat(@"  ,[FURANACEUP1] AS '上爐1-1',[FURANACEUP2] AS '上爐2-1',[FURANACEUP3] AS '上爐3-1',[FURANACEUP4] AS '上爐4-1',[FURANACEUP5] AS '上爐5-1'");
                STR.AppendFormat(@"  ,[FURANACEUP1A] AS '上爐1-2',[FURANACEUP2A] AS '上爐2-2',[FURANACEUP3A] AS '上爐3-2',[FURANACEUP4A] AS '上爐4-2',[FURANACEUP5A] AS '上爐5-2'");
                STR.AppendFormat(@"  ,[FURANACEUP1B] AS '上爐1-3',[FURANACEUP2B] AS '上爐2-3',[FURANACEUP3B] AS '上爐3-3',[FURANACEUP4B] AS '上爐4-3',[FURANACEUP5B] AS '上爐5-3'");
                STR.AppendFormat(@"  ,[FURANACEDOWN1] AS '下爐1-1',[FURANACEDOWN2] AS '下爐2-1',[FURANACEDOWN3] AS '下爐3-1',[FURANACEDOWN4] AS '下爐4-1',[FURANACEDOWN5] AS '下爐5-1'");
                STR.AppendFormat(@"  ,[FURANACEDOWN1A] AS '下爐1-2',[FURANACEDOWN2A] AS '下爐2-2',[FURANACEDOWN3A] AS '下爐3-2',[FURANACEDOWN4A] AS '下爐4-2',[FURANACEDOWN5A] AS '下爐5-2'");
                STR.AppendFormat(@"  ,[FURANACEDOWN1B] AS '下爐1-3',[FURANACEDOWN2B] AS '下爐2-3',[FURANACEDOWN3B] AS '下爐3-3',[FURANACEDOWN4B] AS '下爐4-3',[FURANACEDOWN5B] AS '下爐5-3'");
                STR.AppendFormat(@"  ,[CHECKOVENMD].[MAIN] AS '線別',CONVERT(varchar(100),[CHECKOVENMD].[MAINDATE], 112)  AS '日期',[CHECKOVENMD].[TARGETPROTA001] AS '單別',[CHECKOVENMD].[TARGETPROTA002] AS '單號',[CHECKOVENMD].[MB001] AS '品號'");
                STR.AppendFormat(@"  ,[CHECKOVENM].[MB002] AS '品名',CONVERT(varchar(100),[CHECKOVENM].[STIME], 8) AS '開始時間',CONVERT(varchar(100),[CHECKOVENM].[ETIME], 8)  AS '結束時間'");
                STR.AppendFormat(@"  ,[CHECKOVENM].[GAS] AS '瓦斯磅數',[CHECKOVENM].[FLODPEOPLE1]  AS '折疊人員1',[CHECKOVENM].[FLODPEOPLE2]   AS '折疊人員2'");
                STR.AppendFormat(@"  ,[CHECKOVENM].[MANAGER]  AS '主管',[CHECKOVENM].[OPERATOR]  AS '操作人員'");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKOVENM] WITH(NOLOCK),[TKCIM].[dbo].[CHECKOVENMD] WITH(NOLOCK)");
                STR.AppendFormat(@"  WHERE [CHECKOVENM].[TARGETPROTA001]=[CHECKOVENMD].[TARGETPROTA001] AND [CHECKOVENM].[TARGETPROTA002]=[CHECKOVENMD].[TARGETPROTA002] ");
                STR.AppendFormat(@"  AND CONVERT(varchar(100),[CHECKOVENMD].[MAINDATE],112)>='{0}' AND CONVERT(varchar(100),[CHECKOVENMD].[MAINDATE],112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY [CHECKOVENMD].[MAIN] ,CONVERT(varchar(100),[CHECKOVENMD].[MAINDATE],112),CONVERT(varchar(100),[MANUTIME], 8) ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds14";
            }
            else if (comboBox1.Text.ToString().Equals("首件檢查記錄表-成型"))
            {

                STR.AppendFormat(@"  SELECT  ");
                STR.AppendFormat(@"  [MAIN] AS '組別',CONVERT(varchar(100),[MAINDATE], 112) AS '日期',CONVERT(varchar(100),[MAINTIME],14) AS '時間',[TARGETPROTA001] AS '單別'");
                STR.AppendFormat(@"  ,[TARGETPROTA002] AS '單號',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格'");
                STR.AppendFormat(@"  ,[CHECKNUM] AS '檢查片數',[WEIGHT] AS '平均重量',[LENGTH] AS '平均長度',[TEMPER] AS '環境溫度'");
                STR.AppendFormat(@"  ,[HUMI] AS '環境溼度',[TIME] AS '烤爐時間',[SPEED] AS '烤爐速度',[OVENTEMP] AS '烤爐溫度'");
                STR.AppendFormat(@"  ,[JUDG] AS '口味判定',[METRAILCHECK] AS '原料投入確認',[TEMP] AS '備註'");
                STR.AppendFormat(@"  ,[FJUDG] AS '判定'");
                STR.AppendFormat(@"  ,[OWNER] AS '填表人',[MANAGER] AS '製造主管',[QC] AS '稽核人員'");
                STR.AppendFormat(@"  ,[ID]");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPE]");
                STR.AppendFormat(@"  WHERE CONVERT(varchar(100),[MAINDATE], 112)>='{0}' AND CONVERT(varchar(100),[MAINDATE], 112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY [MAIN] ,CONVERT(varchar(100),[MAINDATE], 112)");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds15";
            }
            else if (comboBox1.Text.ToString().Equals("首件檢查記錄表-成品"))
            {
                STR.AppendFormat(@"  SELECT");
                STR.AppendFormat(@"  [MAIN] AS '組別',CONVERT(varchar(100),[MAINDATE], 112) AS '日期',CONVERT(varchar(100),[MAINTIME],14)  AS '時間',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                STR.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[UNIT] AS '入數單位'");
                STR.AppendFormat(@"  ,[PACKAGENUM] AS '入數數量',[CHECKNUM] AS '抽檢數量',[WEIGHT] AS '重量(公斤/箱)',[TYPEDATE] AS '日期別'");
                STR.AppendFormat(@"  ,[PRODATE] AS '生產/製造日期',[OUTDATE] AS '保質/有效日期',[PACKAGELABEL] AS '外包裝標示',[INLABEL] AS '內容物封口',[TASTEJUDG] AS '口味判定',[TASTEFELL] AS '口感判定',[TEMP] AS '備註'");
                STR.AppendFormat(@"  ,[FJUDG] AS '判定',[OWNER] AS '填表人',[MANAGER] AS '包裝主管',[QC] AS '稽核人員'");
                STR.AppendFormat(@"  ,[ID]");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPEPACKAGE]");
                STR.AppendFormat(@"  WHERE CONVERT(varchar(100),[MAINDATE], 112)>='{0}' AND CONVERT(varchar(100),[MAINDATE], 112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY [MAIN],CONVERT(varchar(100),[MAINDATE], 112)");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds16";
            }
            else if (comboBox1.Text.ToString().Equals("首件檢查記錄表-冷卻"))
            {

                STR.AppendFormat(@"  SELECT ");
                STR.AppendFormat(@"  [MAIN] AS '組別',CONVERT(varchar(100),[MAINDATE], 112) AS '日期',CONVERT(varchar(100),[MAINTIME],14) AS '時間'");
                STR.AppendFormat(@"  ,[CHECKNUM] AS '抽檢數量',[OUTLOOK] AS '色澤外觀',[CHECKFIRSTTYPECOLD].[COOKTEMPER] AS '熟餅溫度(C)'");
                STR.AppendFormat(@"  ,[CHECKFIRSTTYPECOLD].[COOKWEIGHT] AS '熟餅重量(g)',[CHECKFIRSTTYPECOLD].[COOKLENGTH] AS '熟餅長度(cm)',[TEMPER] AS '環境溫度(C)'");
                STR.AppendFormat(@"  ,[HUMI] AS '環境溼度(%)',[TASTEJUDG] AS '口味判定',[TASTEFEEL] AS '口感判定',[TEMP] AS '備註'");
                STR.AppendFormat(@"  ,[FJUDG] AS '判定',[OWNER] AS '填表人',[MANAGER] AS '製造主管',[QC] AS '稽核人員'");
                STR.AppendFormat(@"  ,[CHECKFIRSTTYPECOLDD].[COOKTEMPER] AS '熟餅溫度',[CHECKFIRSTTYPECOLDD].[COOKWEIGHT] AS '熟餅重量'");
                STR.AppendFormat(@"  ,[CHECKFIRSTTYPECOLDD].[COOKLENGTH] AS '熟餅長度',[CHECKFIRSTTYPECOLDD].[MB002]  AS '品名',[CHECKFIRSTTYPECOLDD].[MB003] AS '規格'");
                STR.AppendFormat(@"  ,[CHECKFIRSTTYPECOLDD].[TARGETPROTA001] AS '單別',[CHECKFIRSTTYPECOLDD].[TARGETPROTA002]  AS '單號',[CHECKFIRSTTYPECOLDD].[MB001]  AS '品號'");
                STR.AppendFormat(@"  ,[SERNO] ");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPECOLD],[TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                STR.AppendFormat(@"  WHERE [CHECKFIRSTTYPECOLD].TARGETPROTA001=[CHECKFIRSTTYPECOLDD].TARGETPROTA001 AND [CHECKFIRSTTYPECOLD].TARGETPROTA002=[CHECKFIRSTTYPECOLDD].TARGETPROTA002");
                STR.AppendFormat(@"  WHERE CONVERT(varchar(100),[MAINDATE], 112)>='{0}' AND CONVERT(varchar(100),[MAINDATE], 112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY [MAIN],CONVERT(varchar(100),[MAINDATE], 112),[CHECKFIRSTTYPECOLD].TARGETPROTA001,[CHECKFIRSTTYPECOLD].TARGETPROTA002,SERNO");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds17";
            }
            else if (comboBox1.Text.ToString().Equals("手工生產日報表"))
            {

                STR.AppendFormat(@" SELECT  ");
                STR.AppendFormat(@"  [MAIN] AS '組別',CONVERT(NVARCHAR,[MAINDATE],112) AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號' ");
                STR.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[OILPREIN] AS '油酥/餡-預計投入'");
                STR.AppendFormat(@"  ,[OILACTIN] AS '油酥/餡-實際投入',[WATERPREIN] AS '水麵/皮-預計投入',[WATERACTIN] AS '水麵/皮-實際投入'");
                STR.AppendFormat(@"  ,[TOTALIN] AS '總投入',[CYCLESIDE] AS '可回收邊料',[NG] AS '不良品',[COOKNG] AS '烘烤不良'");
                STR.AppendFormat(@"  ,[OILWORKTIME] AS '油酥/餡-工時',[OILWORKHR] AS '油酥/餡-人數',[WATERWORKTIME] AS '水麵/皮-工時'");
                STR.AppendFormat(@"  ,[WATERWORKHR] AS '水麵/皮-人數',[WORKTIME] AS '製造工時',[WORKHR] AS '製造人數',[CHOREWORK] AS '巧克力-再加工投入'");
                STR.AppendFormat(@"  ,[CHONG] AS '巧克力-不良',[CHOTIME] AS '巧克力-工時',[CHOHR] AS '巧克力-人數',[PACKTIME] AS '後段包裝-工時'");
                STR.AppendFormat(@"  ,[PACKHR] AS '後段包裝-人數',[PACKNG] AS '包裝時餅乾不良',[NGMB002] AS '包裝不良品名',[NGMB003] AS '包裝不良規格'");
                STR.AppendFormat(@"  ,[NGNUM] AS '包裝不良數量',[HALFNUM] AS '半成品數量',[FINALNUM] AS '成品數量',[REMARK] AS '備註'");
                STR.AppendFormat(@"  ,[OWNER] AS '填表人',[ID]");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTHAND]");
                STR.AppendFormat(@"  WHERE CONVERT(varchar(100),[MAINDATE], 112)>='{0}' AND CONVERT(varchar(100),[MAINDATE], 112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY [MAIN],CONVERT(NVARCHAR,[MAINDATE],112),[TARGETPROTA001],[TARGETPROTA002]");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds18";
            }
            else if (comboBox1.Text.ToString().Equals(""))
            {


                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "";
            }



            return STR;
        }

        public void ExcelExport()
        {
            Search();
            string TABLENAME="報表";

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables[tablename];
            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }


            int j = 0;
            if (tablename.Equals("TEMPds1"))
            {
                TABLENAME = "生產日報的分析表";
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(Convert.ToDateTime(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString()).ToString("yyyy/MM/dd"));
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                    ws.GetRow(j + 1).CreateCell(11).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString()));
                    ws.GetRow(j + 1).CreateCell(12).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString()));
                    ws.GetRow(j + 1).CreateCell(13).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString()));
                    ws.GetRow(j + 1).CreateCell(14).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString()));
                    ws.GetRow(j + 1).CreateCell(15).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[15].ToString()));

                    j++;
                }

            }
            else if (tablename.Equals("TEMPds2"))
            {
                TABLENAME = "生產日報的月份分析表";
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());

                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString()));
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                    ws.GetRow(j + 1).CreateCell(11).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString()));
                    ws.GetRow(j + 1).CreateCell(12).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString()));
                    ws.GetRow(j + 1).CreateCell(13).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString()));
                    ws.GetRow(j + 1).CreateCell(14).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString()));


                    j++;
                }

            }
            else if (tablename.Equals("TEMPds3"))
            {
                TABLENAME = "不良品餅乾報廢分析表";
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(Convert.ToDateTime(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString()).ToString("yyyy/MM/dd"));
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString()));
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());

                    j++;
                }

            }

            else if (tablename.Equals("TEMPds4"))
            {
                TABLENAME = "生產日報表明細表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds4"].Rows[i][rows].ToString());
                    }
                }
            }
            else if (tablename.Equals("TEMPds5"))
            {
                TABLENAME = "不良餅麩明細表";
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(Convert.ToDateTime(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString()).ToString("yyyy/MM/dd"));
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());                    
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());

                    j++;
                }
            }
            else if (tablename.Equals("TEMPds6"))
            {
                TABLENAME = "不良邊料明細表";
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(Convert.ToDateTime(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString()).ToString("yyyy/MM/dd"));
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());

                    j++;
                }
            }
            else if (tablename.Equals("TEMPds7"))
            {
                TABLENAME = "不良未熟明細表";
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(Convert.ToDateTime(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString()).ToString("yyyy/MM/dd"));
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());                    
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());

                    j++;
                }
            }
            else if (tablename.Equals("TEMPds8"))
            {
                TABLENAME = "烤爐溫度明細表";

                for(int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds8"].Rows[i][rows].ToString());
                    }
                }
                
            }
            else if (tablename.Equals("TEMPds9"))
            {
                TABLENAME = "成型檢驗表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds9"].Rows[i][rows].ToString());
                    }
                }
            }
            else if (tablename.Equals("TEMPds10"))
            {
                TABLENAME = "水麵添加表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds10"].Rows[i][rows].ToString());
                    }
                }
            }
            else if (tablename.Equals("TEMPds11"))
            {
                TABLENAME = "油酥添加表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds11"].Rows[i][rows].ToString());
                    }
                }
            }
            else if (tablename.Equals("TEMPds12"))
            {
                TABLENAME = "成型檢驗表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds12"].Rows[i][rows].ToString());
                    }
                }
            }
            else if (tablename.Equals("TEMPds13"))
            {
                TABLENAME = "出爐餅溫量測記錄表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds13"].Rows[i][rows].ToString());
                    }
                }
            }
            else if (tablename.Equals("TEMPds14"))
            {
                TABLENAME = "烘烤製程記錄";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds14"].Rows[i][rows].ToString());
                    }
                }
            }
            else if (tablename.Equals("TEMPds15"))
            {
                TABLENAME = "首件檢查記錄表-成型";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds15"].Rows[i][rows].ToString());
                    }
                }
            }
            else if (tablename.Equals("TEMPds16"))
            {
                TABLENAME = "首件檢查記錄表-成品";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds16"].Rows[i][rows].ToString());
                    }
                }
            }
            else if (tablename.Equals("TEMPds17"))
            {               
                TABLENAME = "首件檢查記錄表 - 冷卻";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds17"].Rows[i][rows].ToString());
                    }
                }

            }
            else if (tablename.Equals("TEMPds18"))
            {
                TABLENAME = "手工生產日報表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds18"].Rows[i][rows].ToString());
                    }
                }

            }
            else if (tablename.Equals(""))
            {

            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\{0}-{1}.xlsx", TABLENAME, DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }
        }

        public void SETFASTREPORT()
        {
            if (comboBox2.Text.Equals("包裝組-生產日報表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\包裝組-生產日報表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("烘培檢驗日報表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\烘培檢驗日報表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("包裝班檢驗表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\包裝班檢驗表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("手工生產日報表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\手工生產日報表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("首件檢查記錄表-冷卻"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\首件檢查記錄表-冷卻.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("首件檢查記錄表-成品"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\首件檢查記錄表-成品.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("首件檢查記錄表-成型"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\首件檢查記錄表-成型.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("報廢記錄"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\報廢記錄.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("烘烤製程記錄"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\烘烤製程記錄.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("生產日報的分析表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\生產日報的分析表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("生產日報的月份分析表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\生產日報的月份分析表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyy"));
               
                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("生產日報表明細表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\生產日報表明細表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));

                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("烤爐溫度明細表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\烤爐溫度明細表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));

                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("成型檢驗表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\成型檢驗表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));

                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("水麵添加表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\水麵添加表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));

                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("油酥添加表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\油酥添加表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));

                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("不良餅麩明細表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\不良餅麩明細表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));

                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("不良邊料明細表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\不良邊料明細表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));

                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("不良未熟明細表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\不良未熟明細表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));

                //report1.Load(@"REPORT\TEST1.frx");

            }
            else if (comboBox2.Text.Equals("不良品餅乾報廢明細表"))
            {
                report1 = new Report();
                report1.Load(@"REPORT\不良品餅乾報廢明細表.frx");

                report1.SetParameterValue("P1", dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker4.Value.ToString("yyyyMMdd"));

                //report1.Load(@"REPORT\TEST1.frx");

            }



            report1.Preview = previewControl1;
            report1.Show();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion


    }
}
