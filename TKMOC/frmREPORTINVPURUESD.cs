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
    public partial class frmREPORTINVPURUESD : Form
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
        SqlTransaction tran;
        int result;

        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        string tablename = null;

        string MD003;
        int rowIndexDG1 = -1;
        int rowIndexDG2 = -1;


        public frmREPORTINVPURUESD()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCHMOCMANULINE(string SDay,string EDay)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MD003 AS '品號',MD035 AS '品名'");
                sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                sbSql.AppendFormat(@"  AND MC001=MD001");
                sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}' ",SDay,EDay);
                sbSql.AppendFormat(@"  GROUP BY MD003,MD035");
                sbSql.AppendFormat(@"  ORDER BY MD003,MD035");
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

                    SEARCHINVPURMOC(MD003,dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));


                }
                else
                {
                    MD003 = null;
                }
            }
            
        }

        public void SEARCHINVPURMOC(string MD003,string SDay,string EDay)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                //查每日原物料的使用量、進貨量、庫存量
                //另外查詢領料數量
                //sbSql.AppendFormat(@"  ,ISNULL((SELECT SUM(TE005) FROM [TK].dbo.MOCTE WHERE TE011+TE012 IN ( SELECT TA001+TA002 FROM [TK].dbo.MOCTA WHERE TE004=TEMP2.MD003 AND TA026=TEMP2.COPTD001 AND TA027=TEMP2.COPTD002 AND TA028=TEMP2.COPTD003 )),0)*-1 AS '領料數量' ");

                sbSql.AppendFormat(@"  SELECT SUM(TEMP4.TNUM) AS '預計庫存量',TEMP2.ID AS '列數',TEMP2.MANU AS '線別',TEMP2.MANUDATE AS '日期',TEMP2.MD003 AS '品號',TEMP2.MD035 AS '品名',TEMP2.TNUM AS '用量'");
               
                sbSql.AppendFormat(@"  ,TEMP2.MB004 AS '單位',TEMP2.MB001 AS '成品',TEMP2.MB002 AS '成品名',TEMP2.PACKAGE AS '成品數',TEMP2.COPTD001 AS '訂單單別',TEMP2.COPTD002 AS '訂單單號',TEMP2.COPTD003 AS '訂單序號' ");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  SELECT ROW_NUMBER() OVER (ORDER BY TEMP.MANUDATE) AS ID,MANU,MANUDATE,MD003,MD035,TNUM,MB004,MB001,MB002,PACKAGE,COPTD001,COPTD002,COPTD003");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  SELECT [MANU],CONVERT(NVARCHAR,[MANUDATE],112) AS MANUDATE,[MD003],[MD035],CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008])))*-1 AS TNUM,[MB004],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[PACKAGE],[COPTD001],[COPTD002],[COPTD003]");
                sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                sbSql.AppendFormat(@"  AND MC001=MD001");
                sbSql.AppendFormat(@"  AND [MANU]='新廠包裝線'");
                sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'",SDay,EDay);
                sbSql.AppendFormat(@"  AND [MD003]='{0}'",MD003);
                sbSql.AppendFormat(@"  UNION ");
                sbSql.AppendFormat(@"  SELECT [MANU],CONVERT(NVARCHAR,[MANUDATE],112) AS MANUDATE,[MD003],[MD035],CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008])))*-1 AS TNUM,[MB004],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[NUM],[COPTD001],[COPTD002],[COPTD003]");
                sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                sbSql.AppendFormat(@"  AND MC001=MD001");
                sbSql.AppendFormat(@"  AND [MANU] NOT IN ('新廠包裝線')");
                sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'", SDay, EDay);
                sbSql.AppendFormat(@"  AND [MD003]='{0}'", MD003);
                sbSql.AppendFormat(@"  UNION");
                sbSql.AppendFormat(@"  SELECT '1進貨',TD012,TD004,MB002,CONVERT(DECIMAL(14,2),(CASE WHEN ISNULL(MD002,'')<>'' THEN (ISNULL(TD008-TD015,0)*MD004/MD003) ELSE (TD008-TD015) END )) ,MB004,NULL,NULL,NULL,TD001,TD002,TD003");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.PURTC,[TK].dbo.PURTD ");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD009  ");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002 AND TD004=MB001 AND TD018='Y' AND TD016='N'");
                sbSql.AppendFormat(@"  AND TD012>='{0}' AND TD012<='{1}' ",SDay,EDay);
                sbSql.AppendFormat(@"  AND TD004='{0}'", MD003);             
                sbSql.AppendFormat(@"  UNION ");
                sbSql.AppendFormat(@"  SELECT '0庫存' AS MANU,CONVERT(NVARCHAR,GETDATE(),112) AS MANUDATE,LA001 AS MD003,MB002,SUM(LA005*LA011) TNUM, MB004,NULL AS MB001,NULL AS MB002,NULL AS PACKAGE,NULL AS COPTD001,NULL AS COPTD002,NULL AS COPTD002");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE LA001=MB001");
                sbSql.AppendFormat(@"  AND  LA009 IN ('20004','20006' )");
                sbSql.AppendFormat(@"  AND LA001='{0}' ",MD003);
                sbSql.AppendFormat(@"  GROUP BY LA001,MB002,MB004");
                sbSql.AppendFormat(@"  UNION");
                sbSql.AppendFormat(@"  SELECT '1手動進出貨',CONVERT(NVARCHAR,INVPURUESD.DATES,112),INVPURUESD.MB001,MB002,NUM ,MB004,NULL,NULL,NULL,NULL,NULL,NULL");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TKMOC].dbo.INVPURUESD ");
                sbSql.AppendFormat(@"  WHERE INVMB.MB001=INVPURUESD.MB001");
                sbSql.AppendFormat(@"  AND INVPURUESD.DATES>='{0}' AND INVPURUESD.DATES<='{1}'", SDay, EDay);
                sbSql.AppendFormat(@"  AND INVPURUESD.MB001='{0}'", MD003);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ) AS TEMP ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ) AS TEMP2 JOIN ");
                sbSql.AppendFormat(@"  (SELECT ROW_NUMBER() OVER (ORDER BY TEMP3.MANUDATE) AS ID,MANU,MANUDATE,MD003,MD035,TNUM,MB004,MB001,MB002,PACKAGE,COPTD001,COPTD002,COPTD003");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  SELECT [MANU],CONVERT(NVARCHAR,[MANUDATE],112) AS MANUDATE,[MD003],[MD035],CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008])))*-1 AS TNUM,[MB004],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[PACKAGE],[COPTD001],[COPTD002],[COPTD003]");
                sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                sbSql.AppendFormat(@"  AND MC001=MD001");
                sbSql.AppendFormat(@"  AND [MANU]='新廠包裝線'");
                sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'", SDay, EDay);
                sbSql.AppendFormat(@"  AND [MD003]='{0}'", MD003);
                sbSql.AppendFormat(@"  UNION ");
                sbSql.AppendFormat(@"  SELECT [MANU],CONVERT(NVARCHAR,[MANUDATE],112) AS MANUDATE,[MD003],[MD035],CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008])))*-1 AS TNUM,[MB004],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[NUM],[COPTD001],[COPTD002],[COPTD003]");
                sbSql.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003");
                sbSql.AppendFormat(@"  WHERE [MOCMANULINE].MB001=MC001");
                sbSql.AppendFormat(@"  AND MC001=MD001");
                sbSql.AppendFormat(@"  AND [MANU] NOT IN ('新廠包裝線')");
                sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'", SDay, EDay);
                sbSql.AppendFormat(@"  AND [MD003]='{0}'", MD003);
                sbSql.AppendFormat(@"  UNION");
                sbSql.AppendFormat(@"  SELECT '1進貨',TD012,TD004,MB002,CONVERT(DECIMAL(14,2),(CASE WHEN ISNULL(MD002,'')<>'' THEN (ISNULL(TD008-TD015,0)*MD004/MD003) ELSE (TD008-TD015) END )) ,MB004,NULL,NULL,NULL,TD001,TD002,TD003");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.PURTC,[TK].dbo.PURTD ");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD009 ");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002 AND TD004=MB001 AND TD018='Y' AND TD016='N'");
                sbSql.AppendFormat(@"  AND TD012>='{0}' AND TD012<='{1}' ", SDay, EDay);
                sbSql.AppendFormat(@"  AND TD004='{0}'",MD003);
                sbSql.AppendFormat(@"  UNION ");
                sbSql.AppendFormat(@"  SELECT '0庫存' AS MANU,CONVERT(NVARCHAR,GETDATE(),112) AS MANUDATE,LA001 AS MD003,MB002,SUM(LA005*LA011) TNUM, MB004,NULL AS MB001,NULL AS MB002,NULL AS PACKAGE,NULL AS COPTD001,NULL AS COPTD002,NULL AS COPTD002");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE LA001=MB001");
                sbSql.AppendFormat(@"  AND  LA009 IN ('20004','20006' )  ");
                sbSql.AppendFormat(@"  AND LA001='{0}' ", MD003);
                sbSql.AppendFormat(@"  GROUP BY LA001,MB002,MB004");
                sbSql.AppendFormat(@"  UNION");
                sbSql.AppendFormat(@"  SELECT '1手動進出貨',CONVERT(NVARCHAR,INVPURUESD.DATES,112),INVPURUESD.MB001,MB002,NUM ,MB004,NULL,NULL,NULL,NULL,NULL,NULL");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TKMOC].dbo.INVPURUESD ");
                sbSql.AppendFormat(@"  WHERE INVMB.MB001=INVPURUESD.MB001");
                sbSql.AppendFormat(@"  AND INVPURUESD.DATES>='{0}' AND INVPURUESD.DATES<='{1}'", SDay, EDay);
                sbSql.AppendFormat(@"  AND INVPURUESD.MB001='{0}'", MD003);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ) AS TEMP3");
                sbSql.AppendFormat(@"  ) AS TEMP4 ON TEMP2.ID>=TEMP4.ID");
                sbSql.AppendFormat(@"  GROUP BY TEMP2.ID,TEMP2.MANU,TEMP2.MANUDATE,TEMP2.MD003,TEMP2.MD035,TEMP2.TNUM,TEMP2.MB004,TEMP2.MB001,TEMP2.MB002,TEMP2.PACKAGE,TEMP2.COPTD001,TEMP2.COPTD002,TEMP2.COPTD003");
                sbSql.AppendFormat(@"  ORDER BY TEMP2.MANUDATE, TEMP2.MANU");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
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

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView2.Rows)
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

        public void ADDINVPURUESD(string KIND,string DATES,string MB001,string NUM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                //sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[INVPURUESD]");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[INVPURUESD]");
                sbSql.AppendFormat(" ([KIND],[DATES],[MB001],[NUM])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')",KIND,DATES,MB001,NUM);
                sbSql.AppendFormat(" ");
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
        
        public void DELINVPURUESD(string MB001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[INVPURUESD]");               
                sbSql.AppendFormat(" WHERE MB001='{0}'",MB001);
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

        public void SEARCHCOPTD(string TD002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();    
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT TD004 AS '品號',TD005  AS '品名',TD006  AS '規格'
                                    FROM [TK].dbo.COPTD
                                    WHERE TD002='{0}'
                                    ORDER BY TD004
                                    ", TD002);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds1.Tables["ds1"];
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

        public void SEARCHBOM(string MD001)
        {

            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    ;WITH BOMOrder AS (

                                    SELECT MD001, MD003 ,MD006,MD007,MD008 ,1 AS BOMLevel,CONVERT(DECIMAL(16,5),1*MD006/MD007) AS NUM,MC004
                                    FROM [TK].dbo.VBOMMD 
                                    UNION ALL	
                                    SELECT A.MD001, B.MD003,B.MD006,B.MD007,B.MD008, (B.BOMLevel + 1) AS BOMLevel,CONVERT(DECIMAL(16,5),B.NUM*A.MD006/A.MD007) AS NUM,B.MC004
                                    FROM [TK].dbo.VBOMMD A
                                    INNER JOIN BOMOrder B ON A.MD003 = B.MD001
                                    )
                                    SELECT  MD003 AS '品號' ,MB002  AS '品名',MD001  AS '成品號'
                                    FROM BOMOrder,[TK].dbo.INVMB
                                    WHERE BOMLevel<=5
                                    AND MD003=MB001
                                    AND MD001='{0}'
                                    GROUP BY MD001, MD003 ,MB002
                                    ORDER BY MD003
                                    ", MD001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds1.Tables["ds1"];
                        dataGridView4.AutoResizeColumns();
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
            SEARCHMOCMANULINE(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
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
        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text)&& !string.IsNullOrEmpty(MD003))
            {
                ADDINVPURUESD("1手動進貨", dateTimePicker3.Value.ToString("yyyy/MM/dd"),MD003, textBox2.Text);
            }

            SEARCHINVPURMOC(MD003, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show(MD003+" 要刪除了?", MD003+ " 要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELINVPURUESD(MD003);
                SEARCHINVPURMOC(MD003, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SEARCHCOPTD(textBox3.Text.Trim());
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SEARCHBOM(textBox4.Text.Trim());
        }

        #endregion


    }
}
