﻿using System;
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
using TKITDLL;


namespace TKMOC
{
    public partial class frmCOPPRE : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
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
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();

        int result;

        string ADDTC001002003;
        string ADDTD001;
        string ADDTD002;
        string ADDTD003;
        string DELTC001002003;
        string DELTD001;
        string DELTD002;
        string DELTD003;

        string STATUSPREMANU = null;
        string STATUSPREINVMBMANU = null;

        List<ADDITEM> ADDTARGET = new List<ADDITEM>();

        public class ADDITEM
        {
            public string ORDERNO { get; set; }
            public string MB001 { get; set; }
            public string MB002 { get; set; }
            public int AMOUNT { get; set; }
            public string  UNIT { get; set; }
            public int PRIORITYS { get; set; }
            public string MANU { get; set; }
            public decimal TIMES { get; set; }
            public int HRS { get; set; }
            public string WDT { get; set; }
            public int WHRS { get; set; }
            public int WSHRS { get; set; }
            public int WEHRS { get; set; }
        }
        public frmCOPPRE()
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
        }


        #region FUNCTION
        public void combobox1load()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT [MANU] FROM [TKMOC].[dbo].[PREMANU] ORDER BY [ID]";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();
            
            dt.Columns.Add("MANU", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MANU";
            comboBox1.DisplayMember = "MANU";
            sqlConn.Close();

        }

        public void combobox2load()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT [MANU] FROM [TKMOC].[dbo].[PREMANU] ORDER BY [ID]";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MANU", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MANU";
            comboBox2.DisplayMember = "MANU";
            sqlConn.Close();

        }

        public void PRESCHEDULE()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [PREORDER].[ORDERNO],[PREORDER].[MB001],[PREORDER].[MB002],[PREORDER].[AMOUNT],[PREORDER].[UNIT],[PREORDER].[PRIORITYS],[PREINVMBMANU].MANU,[PREINVMBMANU].TIMES");
                sbSql.AppendFormat(@"  ,CONVERT(INT,ROUND([PREORDER].[AMOUNT]/[PREINVMBMANU].TIMES,0)) AS HRS");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PREORDER],[TKMOC].[dbo].[PREINVMBMANU]");
                sbSql.AppendFormat(@"  WHERE [PREORDER].MB001=[PREINVMBMANU].MB001");
                sbSql.AppendFormat(@"  ORDER BY [PREINVMBMANU].MANU,[PREORDER].[PRIORITYS] DESC,[ORDERNO]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        ADDNEWTARGET(ds1.Tables["TEMPds1"]);

                        ////dataGridView1.Rows.Clear();
                        //dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        //dataGridView1.AutoResizeColumns();
                        ////dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        //先用sql把訂單、優先權、線別排出順序
        //再用工時2小時去排序每張訂單的預計開工日、時間
        //每張訂單的總時數要控制 TWORKHRS、CHECKHRS
        //每天不得超過16小時。超過就跨天 WORKHRS
        //星期日不排 DayOfWeek.Sunday
        public void ADDNEWTARGET(DataTable dt)
        {
            ADDTARGET.Clear();

            int WORKHRS = 0;    //每日工時累積
            int TWORKHRS = 0;   //每張訂單的總工時
            double CHECKHRS = 0;
            int day = 0;    //日期是否跨日
            int WSHRS = 8;  //開工時間
            int WEHRS = 8;  //結束時間

            double LIMITEHRS = 15.9;

            DateTime wdt = new DateTime();
            wdt = dateTimePicker1.Value;

            if (wdt.DayOfWeek==DayOfWeek.Sunday)
            {
                wdt = wdt.AddDays(1);
            }

            bool result;
            string LASTMANU = "START";

            foreach (DataRow od in dt.Rows)
            {
                TWORKHRS = 0;
             
                CHECKHRS = (Convert.ToDouble(od["HRS"].ToString())-0.1);

                //MessageBox.Show(LASTMANU+" "+ od["MANU"].ToString());

                if (!LASTMANU.Equals(od["MANU"].ToString()))
                {
                    WORKHRS = 0;                  
                    TWORKHRS = 0;
                    day = 0;

                    wdt = dateTimePicker1.Value;
                    //改成迴圈檢查行事曆
                    if (wdt.DayOfWeek == DayOfWeek.Sunday)
                    {
                        wdt = wdt.AddDays(1);
                    }

                    WSHRS = 8;
                    WEHRS = 8;

                    //MessageBox.Show(LASTMANU + " " + od["MANU"].ToString());
                }

                while ((CHECKHRS) >= TWORKHRS)
                {
                    if (WORKHRS <= LIMITEHRS)
                    {
                        ADDTARGET.Add(new ADDITEM { ORDERNO = od["ORDERNO"].ToString(), MB001 = od["MB001"].ToString(), MB002 = od["MB002"].ToString(), AMOUNT = Convert.ToInt16(od["AMOUNT"].ToString()), UNIT = od["UNIT"].ToString(), PRIORITYS = Convert.ToInt16(od["PRIORITYS"].ToString()), MANU = od["MANU"].ToString(), TIMES = Convert.ToDecimal(od["TIMES"].ToString()), HRS = Convert.ToInt16(od["HRS"].ToString()), WDT = wdt.ToString("yyyyMMdd"), WHRS = (WORKHRS+2 ), WSHRS = WSHRS , WEHRS = WEHRS+2 });
                        WORKHRS = WORKHRS + 2;
                        WSHRS = WSHRS + 2;
                        WEHRS = WEHRS + 2;
                    }
                    else if (WORKHRS > LIMITEHRS)
                    {
                        WORKHRS = 0;
                        WSHRS = 8;
                        WEHRS = 8;
                        //day = day + 1;
                        wdt = wdt.AddDays(1);

                        if (wdt.DayOfWeek == DayOfWeek.Sunday)
                        {
                            wdt = wdt.AddDays(1);
                        }

                        ADDTARGET.Add(new ADDITEM { ORDERNO = od["ORDERNO"].ToString(), MB001 = od["MB001"].ToString(), MB002 = od["MB002"].ToString(), AMOUNT = Convert.ToInt16(od["AMOUNT"].ToString()), UNIT = od["UNIT"].ToString(), PRIORITYS = Convert.ToInt16(od["PRIORITYS"].ToString()), MANU = od["MANU"].ToString(), TIMES = Convert.ToDecimal(od["TIMES"].ToString()), HRS = Convert.ToInt16(od["HRS"].ToString()), WDT = wdt.ToString("yyyyMMdd"), WHRS = (WORKHRS+2 ), WSHRS = WSHRS, WEHRS = WEHRS + 2 });
                        WORKHRS = WORKHRS + 2;
                        WSHRS = WSHRS + 2;
                        WEHRS = WEHRS + 2;
                    }

                    TWORKHRS = TWORKHRS + 2;
                }




                LASTMANU = od["MANU"].ToString();
            }


            var bindingList = new BindingList<ADDITEM>(ADDTARGET);
            var source = new BindingSource(bindingList, null);
            dataGridView1.DataSource = source;
            dataGridView1.AutoResizeColumns();
        }

        public void SEARCHPREORDER()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT [ORDERNO] AS '訂單',[CLIENT] AS '客戶',[OUTDATES] AS '預交日',[MB001] AS '品號',[MB002] AS '品名',[AMOUNT] AS '數量',[UNIT] AS '單位',[PRIORITYS] AS '優先權'");
                sbSql.AppendFormat(@" FROM [TKMOC].[dbo].[PREORDER] ");
                sbSql.AppendFormat(@" ORDER BY [CLIENT],[OUTDATES] ");
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
                sqlConn.Close();
            }
        }

        public void SEARCHERPCOPTD()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TD001 AS '訂單',TD002  AS '訂單號',TD003  AS '訂單序',TC053 AS '客戶',TD013 AS '預交日',TD004 AS '品號',TD005 AS '品名',CASE WHEN ISNULL(MD001,'')<>'' THEN (TD008+TD024-TD009-TD025)*MD004 ELSE (TD008+TD024-TD009-TD025) END  AS '數量',MB004 AS '單位'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMD] ON MD001=TD004 AND MD002=TD010");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=TD004");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD021='Y' AND TD016='N' AND COPTD.UDF01='Y' ");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND (TD008+TD024-TD009-TD025)>0");            
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TD001+'-'+TD002+'-'+TD003 NOT IN (SELECT [ORDERNO] FROM [TKMOC].[dbo].[PREORDER])");
                sbSql.AppendFormat(@"  ORDER BY TC053,TD013");
                sbSql.AppendFormat(@"  ");
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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            DELTC001002003 = null;
            

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    DELTC001002003 = row.Cells["訂單"].Value.ToString();
                   

                }
                else
                {
                    DELTC001002003 = null;
                  

                }
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            ADDTD001 = null;
            ADDTD002 = null;
            ADDTD003 = null;


            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    ADDTD001 = row.Cells["訂單"].Value.ToString();
                    ADDTD002 = row.Cells["訂單號"].Value.ToString();
                    ADDTD003 = row.Cells["訂單序"].Value.ToString();

                }
                else
                {
                    ADDTD001 = null;
                    ADDTD002 = null;
                    ADDTD003 = null;

                }
            }
        }

        public void ADDPREORDER(string ADDT2001,string ADDTD002,string ADDTD003)
        {
            if (!string.IsNullOrEmpty(ADDT2001)&& !string.IsNullOrEmpty(ADDTD002) && !string.IsNullOrEmpty(ADDTD003))
            {
                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO  [TKMOC].[dbo].[PREORDER]");
                    sbSql.AppendFormat(" ([ORDERNO],[CLIENT],[OUTDATES],[MB001],[MB002],[AMOUNT],[UNIT],[PRIORITYS])");
                    sbSql.AppendFormat(" SELECT TD001+'-'+TD002+'-'+TD003,TC053,TD013,TD004,TD005,CASE WHEN ISNULL(MD001,'')<>'' THEN (TD008+TD024-TD009-TD025)*MD004 ELSE (TD008+TD024-TD009-TD025) END,MB004,'1'");
                    sbSql.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                    sbSql.AppendFormat(" LEFT JOIN [TK].dbo.[INVMD] ON MD001=TD004 AND MD002=TD010");
                    sbSql.AppendFormat(" LEFT JOIN [TK].dbo.[INVMB] ON MB001=TD004");
                    sbSql.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                    sbSql.AppendFormat(" AND TD001+TD002+TD003='{0}'", ADDTD001+ ADDTD002+ ADDTD003);
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

                        MessageBox.Show("完成");
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
        }

        public void DELPREORDER(string DELTC001002003)
        {
            if (!string.IsNullOrEmpty(DELTC001002003))
            {
                try
                {

                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[PREORDER] WHERE [ORDERNO]='{0}'", DELTC001002003);
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

                        MessageBox.Show("完成");
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
        }

        public void UPDATEPREORDER(string ORDERNO, string PRIORITYS)
        {
            if (!string.IsNullOrEmpty(ORDERNO) && !string.IsNullOrEmpty(PRIORITYS))
            {
                try
                {

                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[PREORDER]");
                    sbSql.AppendFormat(" SET [PRIORITYS]='{0}'", PRIORITYS);
                    sbSql.AppendFormat(" WHERE [ORDERNO]='{0}'", ORDERNO);
                    sbSql.AppendFormat(" ");
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

                        MessageBox.Show("完成");
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
        }

        public void SEARCHPREINVMB()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格'");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PREINVMB] ");
                sbSql.AppendFormat(@"  ORDER BY [MB001]");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {

                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds4.Tables["ds4"];
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

        public void ADDPREINVMB()
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO  [TKMOC].[dbo].[PREINVMB]");
                sbSql.AppendFormat(" ([MB001],[MB002],[MB003])");
                sbSql.AppendFormat(" SELECT MB001,MB002,MB003");
                sbSql.AppendFormat(" FROM [TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE (MB001 LIKE '3%' OR MB001 LIKE '4%')");
                sbSql.AppendFormat(" AND MB002 NOT LIKE '%停%'");
                sbSql.AppendFormat(" AND [MB001] NOT IN (SELECT [MB001] FROM [TKMOC].[dbo].[PREINVMB])");
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

                    MessageBox.Show("完成");
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

        public void SERACHPREMANU()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT [ID] AS '代號',[MANU] AS '線別' FROM [TKMOC].[dbo].[PREMANU] ");
                sbSql.AppendFormat(@"  ");

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "ds5");
                sqlConn.Close();


                if (ds5.Tables["ds5"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds5.Tables["ds5"].Rows.Count >= 1)
                    {

                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds5.Tables["ds5"];
                        dataGridView5.AutoResizeColumns();
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
        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;

            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    textBox1.Text = row.Cells["代號"].Value.ToString();
                    textBox2.Text = row.Cells["線別"].Value.ToString();

                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;

                }
            }
        }

        public void SETSTATUS()
        {
            textBox1.Text = null;
            textBox2.Text = null;

            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;


        }
        public void SETSTATUS2()
        {
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
        }

        public void SETSTAUSFIANL()
        {
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;

        }

        public void SETSTATUS3()
        {
            textBox9.ReadOnly = false;
           
        }

        public void SETSTAUSFIANL2()
        {
            textBox9.ReadOnly = true;          

        }

        public void UPDATEPREMANU(string ID,string MANU)
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[PREMANU]");
                sbSql.AppendFormat(" SET [MANU]='{0}'",MANU);
                sbSql.AppendFormat(" WHERE [ID]='{0}'",ID);
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

        public void ADDPREMANU(string ID, string MANU)
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[PREMANU]");
                sbSql.AppendFormat(" ([ID],[MANU])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}')",ID,MANU);
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

        public void DELPREMANU(string ID)
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[PREMANU] ");
                sbSql.AppendFormat(" WHERE [ID]='{0}'",ID);
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


        public void SEARCHPREINVMB2()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if (!string.IsNullOrEmpty(textBox8.Text))
                {
                    sbSql.AppendFormat(@"  SELECT [MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格'");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PREINVMB] ");
                    sbSql.AppendFormat(@"  WHERE [MB001] LIKE '{0}%' ", textBox8.Text);
                    sbSql.AppendFormat(@"  ORDER BY [MB001]");
                    sbSql.AppendFormat(@"  ");
                }
             

                adapter6 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder6 = new SqlCommandBuilder(adapter6);
                sqlConn.Open();
                ds6.Clear();
                adapter6.Fill(ds6, "ds6");
                sqlConn.Close();


                if (ds6.Tables["ds6"].Rows.Count == 0)
                {
                    dataGridView6.DataSource = null;
                }
                else
                {
                    if (ds6.Tables["ds6"].Rows.Count >= 1)
                    {

                        //dataGridView1.Rows.Clear();
                        dataGridView6.DataSource = ds6.Tables["ds6"];
                        dataGridView6.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            textBox3.Text = null;
            textBox5.Text = null;

            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
                    textBox3.Text = row.Cells["品號"].Value.ToString();
                    textBox5.Text = row.Cells["品名"].Value.ToString();

                    SEARCHPREINVMBMANU(textBox3.Text);
                }
                else
                {
                    textBox3.Text = null;
                    textBox5.Text = null;

                }
            }
        }

       

        public void SEARCHPREINVMBMANU(string MB001)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if (!string.IsNullOrEmpty(MB001))
                {
                    sbSql.AppendFormat(@"  SELECT [MB001] AS '品號',[MB002] AS '品名',[MANU] AS '線別',[TIMES] AS '每小時生產量'");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PREINVMBMANU]");
                    sbSql.AppendFormat(@"  WHERE [MB001] LIKE '{0}%' ", MB001.Trim());
                    sbSql.AppendFormat(@"  ");
                }


                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "ds7");
                sqlConn.Close();


                if (ds7.Tables["ds7"].Rows.Count == 0)
                {
                    dataGridView7.DataSource = null;
                }
                else
                {
                    if (ds7.Tables["ds7"].Rows.Count >= 1)
                    {

                        //dataGridView1.Rows.Clear();
                        dataGridView7.DataSource = ds7.Tables["ds7"];
                        dataGridView7.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        public void ADDPREINVMBMANU(string MB001,string MB002,string MANU,decimal TIMES)
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[PREINVMBMANU]");
                sbSql.AppendFormat(" ([MB001],[MB002],[MANU],[TIMES])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')",MB001,MB002,MANU,TIMES);
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

        public void UPDATEPREINVMBMANU(string MB001,string MANU,decimal TIMES)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[PREINVMBMANU]");
                sbSql.AppendFormat(" SET [TIMES]='{0}'", TIMES);
                sbSql.AppendFormat(" WHERE [MB001]='{0}' AND [MANU]='{1}'", MB001, MANU);
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

        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            textBox6.Text = null;
            textBox9.Text = null;
            comboBox2.Text = null;

            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    textBox6.Text = row.Cells["品號"].Value.ToString();
                    textBox9.Text = row.Cells["每小時生產量"].Value.ToString();
                    comboBox2.Text = row.Cells["線別"].Value.ToString();


                }
                else
                {
                    textBox6.Text = null;
                    textBox9.Text = null;
                    comboBox2.Text = null;

                }
            }
        }

        public void DELPREINVMBMANU(string MB001,string MANU)
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[PREINVMBMANU]");              
                sbSql.AppendFormat(" WHERE [MB001]='{0}' AND [MANU]='{1}'", MB001, MANU);
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

        public void SAVETODB()
        {
            StringBuilder SQLINSERT = new StringBuilder();
            SQLINSERT.Clear();


            SQLINSERT.AppendFormat(@" DELETE [TKMOC].[dbo].[PRERESULT]");
            SQLINSERT.AppendFormat(@" ");

            foreach (var obj in ADDTARGET)
            {
                SQLINSERT.AppendFormat(@" INSERT INTO [TKMOC].[dbo].[PRERESULT]");
                SQLINSERT.AppendFormat(@" ([ORDERNO],[MB001],[MB002],[AMOUNT],[UNIT],[PRIORITYS],[MANU],[TIMES],[HRS],[WDT],[WHRS],[WSHRS],[WEHRS])");
                SQLINSERT.AppendFormat(@" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}')",obj.ORDERNO.ToString(),obj.MB001.ToString(),obj.MB002.ToString(),obj.AMOUNT.ToString(),obj.UNIT.ToString(),obj.PRIORITYS.ToString(),obj.MANU.ToString(),obj.TIMES.ToString(),obj.HRS.ToString(),obj.WDT.ToString(),obj.WHRS.ToString(),obj.WSHRS.ToString(),obj.WEHRS.ToString());
                SQLINSERT.AppendFormat(@" ");

            }


            if(!string.IsNullOrEmpty(ADDTARGET.ToString()))
            {
                try
                {

                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(@" DELETE [TKMOC].[dbo].[PRERESULT]");
                    sbSql.AppendFormat(@" {0}", SQLINSERT.ToString());


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

        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1();
            

            Report report1 = new Report();
            report1.Load(@"REPORT\訂單預排報表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT [ORDERNO],[MB001],[MB002],[AMOUNT],[UNIT],[PRIORITYS],[MANU],[TIMES],[HRS],[WDT]");
            SB.AppendFormat(" FROM [TKMOC].[dbo].[PRERESULT]");
            SB.AppendFormat(" GROUP BY [ORDERNO],[MB001],[MB002],[AMOUNT],[UNIT],[PRIORITYS],[MANU],[TIMES],[HRS],[WDT]");
            SB.AppendFormat(" ORDER BY [MANU],[PRIORITYS] DESC,[ORDERNO],[WDT]");
            SB.AppendFormat("  ");

            return SB;
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            PRESCHEDULE();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHPREORDER();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCHERPCOPTD();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADDPREORDER(ADDTD001,ADDTD002,ADDTD003);

            SEARCHPREORDER();
            SEARCHERPCOPTD();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DELPREORDER(DELTC001002003);

            SEARCHPREORDER();
            SEARCHERPCOPTD();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            UPDATEPREORDER(DELTC001002003, numericUpDown1.Value.ToString());

            SEARCHPREORDER();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SEARCHPREINVMB();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ADDPREINVMB();
            SEARCHPREINVMB();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            SERACHPREMANU();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            STATUSPREMANU = "ADD";
            SETSTATUS();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            STATUSPREMANU = "EDIT";
            SETSTATUS2();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (STATUSPREMANU.Equals("EDIT"))
            {
                UPDATEPREMANU(textBox1.Text,textBox2.Text);
            }
            else if (STATUSPREMANU.Equals("ADD"))
            {
                ADDPREMANU(textBox1.Text, textBox2.Text);
            }

            STATUSPREMANU = null;

            SETSTAUSFIANL();
            SERACHPREMANU();
            MessageBox.Show("完成");
        }

        private void button13_Click(object sender, EventArgs e)
        {
            STATUSPREMANU = null;
            string message = " 要刪除了?";

            DialogResult dialogResult = MessageBox.Show(message.ToString(), "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELPREMANU(textBox1.Text);

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SERACHPREMANU();
            MessageBox.Show("完成");

        }
        private void button16_Click(object sender, EventArgs e)
        {
            SEARCHPREINVMB2();
        }


        private void button15_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox3.Text)&& !string.IsNullOrEmpty(textBox5.Text)&& !string.IsNullOrEmpty(comboBox1.Text)&& !string.IsNullOrEmpty(textBox4.Text))
            {
                ADDPREINVMBMANU(textBox3.Text.Trim(), textBox5.Text.Trim(), comboBox1.Text, Convert.ToDecimal(textBox4.Text));

                SEARCHPREINVMBMANU(textBox3.Text);
            }
           
        }
        private void button17_Click(object sender, EventArgs e)
        {
            STATUSPREINVMBMANU = "EDIT";
            SETSTATUS3();
        }
        private void button18_Click(object sender, EventArgs e)
        {
            if (STATUSPREINVMBMANU.Equals("EDIT"))
            {
                UPDATEPREINVMBMANU(textBox6.Text,comboBox2.Text,Convert.ToDecimal(textBox9.Text));
            }

            STATUSPREINVMBMANU = null;

            SETSTAUSFIANL2();
            SEARCHPREINVMBMANU(textBox6.Text);

            MessageBox.Show("完成");
        }

        private void button19_Click(object sender, EventArgs e)
        {

            STATUSPREINVMBMANU = null;
            string message = " 要刪除了?";

            DialogResult dialogResult = MessageBox.Show(message.ToString(), "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELPREINVMBMANU(textBox6.Text, comboBox2.Text);

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHPREINVMBMANU(textBox6.Text);
            MessageBox.Show("完成");

           
        }
        private void button14_Click(object sender, EventArgs e)
        {
            SAVETODB();

            SETFASTREPORT();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            SEARCHPREINVMBMANU(textBox3.Text);
        }
        private void button21_Click(object sender, EventArgs e)
        {
            ADDPREINVMB();
        }

        #endregion


    }
}
