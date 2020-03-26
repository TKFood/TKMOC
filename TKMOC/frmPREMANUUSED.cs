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
    public partial class frmPREMANUUSED : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();

        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        int result;

        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();

        string tablename = null;

        string MD003;
        int rowIndexDG1 = -1;
        int rowIndexDG2 = -1;

        public frmPREMANUUSED()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCHMOCMANULINE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if (comboBox1.Text.Equals("不含水麵")&& comboBox2.Text.Equals("過濾指定品號"))
                {
                    sbSql.AppendFormat(@"  SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',[MB004]  AS '單位'");
                    sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'");
                    sbSql.AppendFormat(@"  ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'");
                    sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '預計採購量'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠包裝線'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製一組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE].[MB001] NOT IN (SELECT [MB001] FROM [TKMOC].[dbo].[PREMANUUSEDINVMB] )");
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製二組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE].[MB001] NOT IN (SELECT [MB001] FROM [TKMOC].[dbo].[PREMANUUSEDINVMB] )");
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製三組(手工)'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE].[MB001] NOT IN (SELECT [MB001] FROM [TKMOC].[dbo].[PREMANUUSEDINVMB] )");
                    sbSql.AppendFormat(@"  ) AS TEMP ");
                    sbSql.AppendFormat(@"  ) AS TEMP2 ");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  GROUP BY [MD003],[MD035],[MB004]");
                    sbSql.AppendFormat(@"  ORDER BY [MD003],[MD035],[MB004]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("不過濾指定品號"))
                {
                    sbSql.AppendFormat(@"  SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',[MB004]  AS '單位'");
                    sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'");
                    sbSql.AppendFormat(@"  ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'");
                    sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '預計採購量'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠包裝線'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製一組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製二組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製三組(手工)'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    
                    sbSql.AppendFormat(@"  ) AS TEMP ");
                    sbSql.AppendFormat(@"  ) AS TEMP2 ");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  GROUP BY [MD003],[MD035],[MB004]");
                    sbSql.AppendFormat(@"  ORDER BY [MD003],[MD035],[MB004]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                //針對製一組、製二組、手工，再往下多查一層bom
                else if(comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("過濾指定品號"))
                {
                    sbSql.AppendFormat(@"  SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',MB004 AS '單位'");
                    sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'");
                    sbSql.AppendFormat(@"  ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'");
                    sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '預計採購量'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],TEMP.[MB001],TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠包裝線'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]");
                    sbSql.AppendFormat(@"  ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC1.MC001");
                    sbSql.AppendFormat(@"  AND MC1.MC001=MD1.MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製一組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE].[MB001] NOT IN (SELECT [MB001] FROM [TKMOC].[dbo].[PREMANUUSEDINVMB] )");
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]");
                    sbSql.AppendFormat(@"  ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC1.MC001");
                    sbSql.AppendFormat(@"  AND MC1.MC001=MD1.MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製二組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE].[MB001] NOT IN (SELECT [MB001] FROM [TKMOC].[dbo].[PREMANUUSEDINVMB] )");
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]");
                    sbSql.AppendFormat(@"  ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC1.MC001");
                    sbSql.AppendFormat(@"  AND MC1.MC001=MD1.MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製三組(手工)'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE].[MB001] NOT IN (SELECT [MB001] FROM [TKMOC].[dbo].[PREMANUUSEDINVMB] )");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ) AS TEMP ");
                    sbSql.AppendFormat(@"  ) AS TEMP2");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]");
                    sbSql.AppendFormat(@"  GROUP BY [MD003],[MD035],[MB004]");
                    sbSql.AppendFormat(@"  ORDER BY [MD003],[MD035],[MB004]");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("不過濾指定品號"))
                {
                    sbSql.AppendFormat(@"  SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',MB004 AS '單位'");
                    sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'");
                    sbSql.AppendFormat(@"  ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'");
                    sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '預計採購量'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],TEMP.[MB001],TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠包裝線'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]");
                    sbSql.AppendFormat(@"  ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC1.MC001");
                    sbSql.AppendFormat(@"  AND MC1.MC001=MD1.MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製一組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]");
                    sbSql.AppendFormat(@"  ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC1.MC001");
                    sbSql.AppendFormat(@"  AND MC1.MC001=MD1.MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製二組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]");
                    sbSql.AppendFormat(@"  ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC1.MC001");
                    sbSql.AppendFormat(@"  AND MC1.MC001=MD1.MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製三組(手工)'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ) AS TEMP ");
                    sbSql.AppendFormat(@"  ) AS TEMP2");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]");
                    sbSql.AppendFormat(@"  GROUP BY [MD003],[MD035],[MB004]");
                    sbSql.AppendFormat(@"  ORDER BY [MD003],[MD035],[MB004]");
                    sbSql.AppendFormat(@"  ");
                }

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["ds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView1.Rows)
                        {
                            //判断
                            if (Convert.ToDecimal(dgRow.Cells[5].Value) > 0)
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.BackColor = Color.Pink;
                            }
                        }

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    SEARCHMOCMANULINE2(MD003);


                }
                else
                {
                    MD003 = null;
                }
            }
        }

        public void SEARCHMOCMANULINE2(string MD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if (comboBox1.Text.Equals("不含水麵"))
                {
                    sbSql.AppendFormat(@"  SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'");
                    sbSql.AppendFormat(@" ,[MB001] AS '成品',[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠包裝線'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製一組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製二組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  ,[MB004]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製三組(手工)'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ) AS TEMP ");
                    sbSql.AppendFormat(@"  WHERE [MD003]='{0}'", MD003);
                    sbSql.AppendFormat(@"  ORDER BY [MANU] ,[MANUDATE],[MD003]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox1.Text.Equals("含水麵"))
                {
                    sbSql.AppendFormat(@"  SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量'");
                    sbSql.AppendFormat(@"  ,[MB001] AS '成品',[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]");
                    sbSql.AppendFormat(@"  ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                    sbSql.AppendFormat(@"  AND MC001=MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠包裝線'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]");
                    sbSql.AppendFormat(@"  ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC1.MC001");
                    sbSql.AppendFormat(@"  AND MC1.MC001=MD1.MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製一組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]");
                    sbSql.AppendFormat(@"  ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC1.MC001");
                    sbSql.AppendFormat(@"  AND MC1.MC001=MD1.MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製二組'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  UNION ALL ");
                    sbSql.AppendFormat(@"  SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]");
                    sbSql.AppendFormat(@"  ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM");
                    sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC1.MC001");
                    sbSql.AppendFormat(@"  AND MC1.MC001=MD1.MD001");
                    sbSql.AppendFormat(@"  AND [MANU]='新廠製三組(手工)'");
                    sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ) AS TEMP ");
                    sbSql.AppendFormat(@"  WHERE [MD003]='{0}'", MD003);
                    sbSql.AppendFormat(@"  ORDER BY [MANU] ,[MANUDATE],[MD003]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }


                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }

        }

        public void SEARCHDG1(string SEARCHSTRING, int INDEX)
        {
            String searchValue = SEARCHSTRING;
            rowIndexDG1 = INDEX;
            int ROWS = 0;

            for (int i = INDEX; i < dataGridView1.Rows.Count; i++)
            {
                ROWS = i;

                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG1 = i;

                    dataGridView1.CurrentRow.Selected = false;
                    dataGridView1.Rows[i].Selected = true;
                    int index = rowIndexDG1;
                    dataGridView1.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
                if (dataGridView1.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG1 = i;

                    dataGridView1.CurrentRow.Selected = false;
                    dataGridView1.Rows[i].Selected = true;
                    int index = rowIndexDG1;
                    dataGridView1.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
            }

            if (ROWS == dataGridView1.Rows.Count - 1)
            {
                if (MessageBox.Show("已查到最後一筆，是否從頭開始?", "已查到最後一筆，是否從頭開始?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    SEARCHDG1(textBox1.Text.Trim(), 0);
                }
                else
                {

                }
            }
        }

        public void SEARCHPREMANUUSEDINVMB()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB001] AS '品號' ,[MB002] AS '品名'");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB]");
                sbSql.AppendFormat(@"  ORDER BY [MB001]");
                sbSql.AppendFormat(@"  ");


                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds3.Tables["ds3"];
                        dataGridView3.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            SEARCHINVMB(textBox2.Text.Trim());
        }

        public void SEARCHINVMB(string MB001)
        {
            textBox3.Text = null;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT MB002,MB004,MB068 ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE MB001='{0}'", MB001);
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        textBox3.Text = ds4.Tables["ds4"].Rows[0]["MB002"].ToString();
                    }
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

        public void ADDPREMANUUSEDINVMB(string MB001,string MB002)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[PREMANUUSEDINVMB]");
                sbSql.AppendFormat(" ([MB001],[MB002])");
                sbSql.AppendFormat(" VALUES('{0}','{1}')",MB001,MB002);
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox4.Text = null;
            textBox5.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox4.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["品名"].Value.ToString().Trim();

                }
                else
                {
                    MD003 = null;
                }
            }
        }
        public void DELETEPREMANUUSEDINVMB(string MB001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[PREMANUUSEDINVMB]");
                sbSql.AppendFormat(" WHERE [MB001]='{0}'", MB001);          
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
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
            SEARCHMOCMANULINE();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            if (rowIndexDG1 == -1)
            {
                SEARCHDG1(textBox1.Text.Trim(), 0);
            }
            else
            {
                SEARCHDG1(textBox1.Text.Trim(), rowIndexDG1 + 1);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SEARCHPREMANUUSEDINVMB();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ADDPREMANUUSEDINVMB(textBox2.Text.Trim(),textBox3.Text.Trim());
            SEARCHPREMANUUSEDINVMB();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETEPREMANUUSEDINVMB(textBox4.Text.Trim());
                SEARCHPREMANUUSEDINVMB();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }



        #endregion


    }
}
