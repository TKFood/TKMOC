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
                        dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];
                        
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
            else if (comboBox1.Text.ToString().Equals("不良品餅乾報廢分析表"))
            {
                STR.AppendFormat(@"  SELECT [MAIN] AS '線別',CONVERT(varchar(100),[MAINDATE], 111) AS '日期',[DAMAGEDCOOKIES] AS '破損餅乾(kg)',[LANDCOOKIES] AS '落地餅乾(kg)',[SCRAPCOOKIES]  AS '餅乾屑(kg)',[ID]");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSCRAPPEDMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'",dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY CONVERT(varchar(100),[MAINDATE], 111),[MAIN] ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds3";
            }
            else if (comboBox1.Text.ToString().Equals("生產日報表明細表"))
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
                STR.AppendFormat(@"  WHERE  [PRODUCEDATE]>='{0}' AND [PRODUCEDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY [PRODUCEDATE],[PRODUCEDEP],[PRODUCENAME]");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds4";
            }
            else if (comboBox1.Text.ToString().Equals("不良餅麩明細表"))
            {
                STR.AppendFormat(@"  SELECT [MAINDATE] AS '日期',CONVERT(varchar(100),[MAINTIME],8)  AS '時間',[MB002] AS '品名',[NUM] AS '回收量',[NGNUM] AS '不良品報廢' ,[MAIN] AS '線別',[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGCOOKIESMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  UNION ALL ");
                STR.AppendFormat(@"  SELECT  '9999/9/9','合計','',SUM([NUM]) AS '回收量',SUM([NGNUM]) AS '不良品報廢' , '','','',''");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGCOOKIESMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY [MAINDATE],[MAIN],[MB001]");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds5";
            }
            else if (comboBox1.Text.ToString().Equals("不良邊料明細表"))
            {
                STR.AppendFormat(@"  SELECT [MAINDATE] AS '日期', CONVERT(varchar(100),[MAINTIME],8)  AS '時間',[MB002] AS '品名',[NUM] AS '回收邊料',[NGNUM] AS '不良品報廢',[MAIN] AS '線別',[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSIDEMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  UNION ALL ");
                STR.AppendFormat(@"  SELECT  '9999/9/9','合計','',SUM([NUM]) AS '回收邊料',SUM([NGNUM]) AS '不良品報廢' , '','','',''");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSIDEMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY [MAINDATE],[MAIN],[MB001]");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds6";
            }
            else if (comboBox1.Text.ToString().Equals("不良未熟明細表"))
            {
                STR.AppendFormat(@"  SELECT  [MAINDATE] AS '日期',CONVERT(varchar(100),[MAINTIME],8) AS '時間',[MB002] AS '品名',[NUM] AS '未熟餅',[COOKTIME] AS '烤培時間',[NGNUM] AS '不良品報廢',[MAIN] AS '線別',[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGNOBURNMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  UNION ALL ");
                STR.AppendFormat(@"  SELECT  '9999/9/9','合計','',SUM([NUM]) AS '未熟餅',SUM([COOKTIME]) AS '烤培時間',SUM([NGNUM]) AS '不良品報廢' , '','','',''");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[NGNOBURNMD]");
                STR.AppendFormat(@"  WHERE [MAINDATE]>='{0}' AND [MAINDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY [MAINDATE],[MAIN],[MB001]");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds7";
            }
            else if (comboBox1.Text.ToString().Equals(""))
            {

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
            else if (tablename.Equals("TEMPds5"))
            {

            }
            else if (tablename.Equals("TEMPds6"))
            {

            }
            else if (tablename.Equals("TEMPds7"))
            {

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
        #endregion


    }
}
