﻿using System;
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

        SqlCommand cmd = new SqlCommand();

        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        string tablename = null;

        string MD003;

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

                sbSql.AppendFormat(@"  SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',[MB004]  AS '單位'");
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
                sbSql.AppendFormat(@"  GROUP BY [MD003],[MD035],[MB004]");
                sbSql.AppendFormat(@"  ORDER BY [MD003],[MD035],[MB004]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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
                sbSql.AppendFormat(@"  AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ) AS TEMP ");
                sbSql.AppendFormat(@"  WHERE [MD003]='{0}'",MD003);
                sbSql.AppendFormat(@"  ORDER BY [MANU] ,[MANUDATE],[MD003]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }

        #endregion

      
    }
}
