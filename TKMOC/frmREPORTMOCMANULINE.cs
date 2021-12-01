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
using System.Threading;
using TKITDLL;

namespace TKMOC
{
    public partial class frmREPORTMOCMANULINE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter22 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder22 = new SqlCommandBuilder();


        SqlDataAdapter adapterCALENDAR = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderCALENDAR = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlDataAdapter adapter6= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlDataAdapter adapter8 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder8 = new SqlCommandBuilder();

        SqlDataAdapter adapter11 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder11 = new SqlCommandBuilder();
        SqlDataAdapter adapter12 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder12 = new SqlCommandBuilder();
        SqlDataAdapter adapter13 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder13 = new SqlCommandBuilder();
        SqlDataAdapter adapter14 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder14 = new SqlCommandBuilder();

        DataSet dsCALENDAR = new DataSet();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();

        DataSet ds2 = new DataSet();
        DataSet ds22 = new DataSet();
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();

        DataSet ds11 = new DataSet();
        DataSet ds12= new DataSet();
        DataSet ds13 = new DataSet();
        DataSet ds14 = new DataSet();

        string tablename = null;
        int rownum = 0;

        string SOURCEID;
        string DATES = null;
        string strDesktopPath;
        string pathFile;
        string pathFile2;
        string pathFile4;

        string[] message = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] message2 = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] message3 = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] message4 = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        DateTime sdt;
        DateTime edt;
        DateTime sdt2;
        DateTime edt2;
        DateTime sdt4;
        DateTime edt4;


        /// <summary>
        /// 製一線桶數 BASELIMITHRSBAR1
        /// 製二線桶數 BASELIMITHRSBAR2
        /// 包裝線稼動率時數 BASELIMITHRS9
        /// 製一線稼動率時數 BASELIMITHRS1
        /// 製二線稼動率時數 BASELIMITHRS2
        /// 手工線稼動率時數 BASELIMITHRS3
        /// </summary>
        decimal BASELIMITHRSBAR1 = 0;
        decimal BASELIMITHRSBAR2 = 0;
        decimal BASELIMITHRS1 = 0;
        decimal BASELIMITHRS2 = 0;
        decimal BASELIMITHRS3 = 0;
        decimal BASELIMITHRS9 = 0;

        public frmREPORTMOCMANULINE()
        {
            InitializeComponent();

            SETCALENDAR();

            //comboBox1load();
            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();
            comboBox5load();
            comboBox6load();
            comboBox7load();

            SETDATE();
            SETDATE2();
        }

        #region FUNCTION
        public void SETDATE()
        {
            DateTime SETDT = Convert.ToDateTime(dateTimePicker9.Value.ToString("yyyy/MM") + "/01");
            DateTime FirstDay = SETDT.AddDays(-SETDT.Day + 1);
            DateTime LastDay = SETDT.AddMonths(1).AddDays(-SETDT.AddMonths(1).Day);
                        
            sdt = FirstDay;
            edt=LastDay;
        }

        public void SETDATE2()
        {
            DateTime SETDT = Convert.ToDateTime(dateTimePicker10.Value.ToString("yyyy/MM") + "/01");
            DateTime FirstDay = SETDT.AddDays(-SETDT.Day + 1);
            DateTime LastDay = SETDT.AddMonths(1).AddDays(-SETDT.AddMonths(1).Day);

            sdt2 = FirstDay;
            edt2 = LastDay;
        }
        public void SETDATE4()
        {
            DateTime SETDT = Convert.ToDateTime(dateTimePicker11.Value.ToString("yyyy/MM") + "/01");
            DateTime FirstDay = SETDT.AddDays(-SETDT.Day + 1);
            DateTime LastDay = SETDT.AddMonths(1).AddDays(-SETDT.AddMonths(1).Day);

            sdt4= FirstDay;
            edt4 = LastDay;
        }
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT '全部' MD001,'全部' MD002 UNION ALL SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE  MD003 IN ('20')    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD002";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox2load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT '全部' MD001,'全部' MD002 UNION ALL SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE  MD003 IN ('20')    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD002";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox3load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT '全部' MD001,'全部' MD002 UNION ALL SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE  MD003 IN ('20')    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "MD002";
            comboBox3.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox4load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT '全部' MD001,'全部' MD002 UNION ALL SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE  MD003 IN ('20')    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "MD002";
            comboBox4.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox5load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT '全部' MD001,'全部' MD002 UNION ALL SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE  MD003 IN ('20')    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "MD002";
            comboBox5.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox6load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE  MD003 IN ('20')    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "MD002";
            comboBox6.DisplayMember = "MD002";
            sqlConn.Close();


        }



        public void comboBox7load()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE  MD003 IN ('20')  UNION ALL SELECT '全部','全部'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "MD002";
            comboBox7.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


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
       

            if (!string.IsNullOrEmpty(comboBox1.Text.ToString())&& !comboBox1.Text.ToString().Equals("全部"))
            {                
                STR.AppendFormat(@"  SELECT [MOCMANULINE].MANU AS '線別',CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) AS '生產日'");
                STR.AppendFormat(@"  ,[MOCMANULINE].MB001 AS '品號',[MOCMANULINE].MB002 AS '品名',[MOCMANULINE].MB003 AS '規格'");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINE].BAR,0) AS '桶數',ISNULL([MOCMANULINE].NUM,0) AS '片數',ISNULL([MOCMANULINE].BOX,0) AS '箱數',ISNULL([MOCMANULINE].PACKAGE,0) AS '包裝數'");
                STR.AppendFormat(@"  ,ISNULL(CONVERT(NVARCHAR(10),[MOCMANULINE].OUTDATE,112),'') AS '預交日'");
                STR.AppendFormat(@"  ,[MOCMANULINE].CLINET AS '客戶',[MOCMANULINE].[TA029] AS '備註',ISNULL([MOCMANULINE].MANUHOUR,0) AS '生產時數'");
                STR.AppendFormat(@"  ,[ID]");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINERESULT].MOCTA001,'') AS '製令單別',ISNULL([MOCMANULINERESULT].MOCTA002,'') AS '製令單號'");
                STR.AppendFormat(@"  ,[MOCTA].TA015 AS '預計產量' ,[MOCTA].TA007 AS '單位'   ");
                STR.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE]");
                STR.AppendFormat(@"  LEFT JOIN  [TKMOC].dbo.[MOCMANULINERESULT] ON [MOCMANULINE].ID=[MOCMANULINERESULT].SID");
                STR.AppendFormat(@"  LEFT JOIN [TK].dbo.[MOCTA] ON [MOCTA].TA001=[MOCMANULINERESULT].MOCTA001 AND [MOCTA].TA002=[MOCMANULINERESULT].MOCTA002");
                STR.AppendFormat(@"  WHERE [MOCMANULINE].MANU='{0}'", comboBox1.Text.ToString());
                STR.AppendFormat(@"  AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112)>='{0}' AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112)<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112),[MOCMANULINE].MB001");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds1";
            }
           
            else if (comboBox1.Text.ToString().Equals("全部"))
            {
                STR.AppendFormat(@"  SELECT [MOCMANULINE].MANU AS '線別',CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) AS '生產日'");
                STR.AppendFormat(@"  ,[MOCMANULINE].MB001 AS '品號',[MOCMANULINE].MB002 AS '品名',[MOCMANULINE].MB003 AS '規格'");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINE].BAR,0) AS '桶數',ISNULL([MOCMANULINE].NUM,0) AS '片數',ISNULL([MOCMANULINE].BOX,0) AS '箱數',ISNULL([MOCMANULINE].PACKAGE,0) AS '包裝數'");
                STR.AppendFormat(@"  ,ISNULL(CONVERT(NVARCHAR(10),[MOCMANULINE].OUTDATE,112),'') AS '預交日'");
                STR.AppendFormat(@"  ,[MOCMANULINE].CLINET AS '客戶',[MOCMANULINE].[TA029] AS '備註',ISNULL([MOCMANULINE].MANUHOUR,0) AS '生產時數'");
                STR.AppendFormat(@"  ,[ID]");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINERESULT].MOCTA001,'') AS '製令單別',ISNULL([MOCMANULINERESULT].MOCTA002,'') AS '製令單號'");
                STR.AppendFormat(@"  ,[MOCTA].TA015 AS '預計產量' ,[MOCTA].TA007 AS '單位'   ");
                STR.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE]");
                STR.AppendFormat(@"  LEFT JOIN  [TKMOC].dbo.[MOCMANULINERESULT] ON [MOCMANULINE].ID=[MOCMANULINERESULT].SID");
                STR.AppendFormat(@"  LEFT JOIN [TK].dbo.[MOCTA] ON [MOCTA].TA001=[MOCMANULINERESULT].MOCTA001 AND [MOCTA].TA002=[MOCMANULINERESULT].MOCTA002");
                STR.AppendFormat(@"  WHERE CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112)>='{0}' AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112)<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112),[MOCMANULINE].MANU,[MOCMANULINE].MB001");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds2";
            }
            



            return STR;
        }

        public void SearchV2()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSqlV2();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


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

                        dataGridView4.DataSource = ds.Tables[tablename];
                        dataGridView4.AutoResizeColumns();
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
        public StringBuilder SETsbSqlV2()
        {
            StringBuilder STR = new StringBuilder();


            if (!string.IsNullOrEmpty(comboBox3.Text.ToString()) && !comboBox3.Text.ToString().Equals("全部"))
            {
                STR.AppendFormat(@"  SELECT [MOCMANULINE].MANU AS '線別',CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) AS '生產日'");
                STR.AppendFormat(@"  ,[MOCMANULINE].MB001 AS '品號',[MOCMANULINE].MB002 AS '品名',[MOCMANULINE].MB003 AS '規格'");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINE].BAR,0) AS '桶數',ISNULL([MOCMANULINE].NUM,0) AS '片數',ISNULL([MOCMANULINE].BOX,0) AS '箱數',ISNULL([MOCMANULINE].PACKAGE,0) AS '包裝數'");
                STR.AppendFormat(@"  ,[MOCMANULINE].CLINET AS '客戶'");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINERESULT].MOCTA001,'') AS '製令單別',ISNULL([MOCMANULINERESULT].MOCTA002,'') AS '製令單號'");
                STR.AppendFormat(@"  ,[MOCTA].TA015 AS '預計產量' ,[MOCTA].TA007 AS '單位'   ");
                STR.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE]");
                STR.AppendFormat(@"  LEFT JOIN  [TKMOC].dbo.[MOCMANULINERESULT] ON [MOCMANULINE].ID=[MOCMANULINERESULT].SID");
                STR.AppendFormat(@"  LEFT JOIN [TK].dbo.[MOCTA] ON [MOCTA].TA001=[MOCMANULINERESULT].MOCTA001 AND [MOCTA].TA002=[MOCMANULINERESULT].MOCTA002");
                STR.AppendFormat(@"  WHERE [MOCMANULINE].MANU='{0}'", comboBox3.Text.ToString());
                STR.AppendFormat(@"  AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112)>='{0}' AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112)<='{1}' ", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112),[MOCMANULINE].MB001");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds3";
            }

            else if (comboBox3.Text.ToString().Equals("全部"))
            {
                STR.AppendFormat(@"  SELECT [MOCMANULINE].MANU AS '線別',CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) AS '生產日'");
                STR.AppendFormat(@"  ,[MOCMANULINE].MB001 AS '品號',[MOCMANULINE].MB002 AS '品名',[MOCMANULINE].MB003 AS '規格'");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINE].BAR,0) AS '桶數',ISNULL([MOCMANULINE].NUM,0) AS '片數',ISNULL([MOCMANULINE].BOX,0) AS '箱數',ISNULL([MOCMANULINE].PACKAGE,0) AS '包裝數'");
                STR.AppendFormat(@"  ,[MOCMANULINE].CLINET AS '客戶'");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINERESULT].MOCTA001,'') AS '製令單別',ISNULL([MOCMANULINERESULT].MOCTA002,'') AS '製令單號'");
                STR.AppendFormat(@"  ,[MOCTA].TA015 AS '預計產量' ,[MOCTA].TA007 AS '單位'   ");
                STR.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE]");
                STR.AppendFormat(@"  LEFT JOIN  [TKMOC].dbo.[MOCMANULINERESULT] ON [MOCMANULINE].ID=[MOCMANULINERESULT].SID");
                STR.AppendFormat(@"  LEFT JOIN [TK].dbo.[MOCTA] ON [MOCTA].TA001=[MOCMANULINERESULT].MOCTA001 AND [MOCTA].TA002=[MOCMANULINERESULT].MOCTA002");
                STR.AppendFormat(@"  WHERE CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112)>='{0}' AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112)<='{1}' ", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112),[MOCMANULINE].MANU,[MOCMANULINE].MB001");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds4";
            }




            return STR;
        }

        public void ExcelExport()
        {
            
            string TABLENAME = "報表";
            int rows = 0;

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

                TABLENAME = "預計製令報表";
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ws.GetRow(j + 1).CreateCell(i).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[i].ToString());
                    }
                    
                    j++;
                }

            }
            else if (tablename.Equals("TEMPds2"))
            {               

                TABLENAME = "預計製令報表";
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ws.GetRow(j + 1).CreateCell(i).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[i].ToString());
                    }

                    j++;
                }

            }
                        

            else if (tablename.Equals("TEMPds3"))
            {
                TABLENAME = "預計製令報表";
                foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                {
                    ws.CreateRow(j + 1);

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ws.GetRow(j + 1).CreateCell(i).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[i].ToString());
                    }

                    j++;
                }
            }
            else if (tablename.Equals("TEMPds4"))
            {
                TABLENAME = "預計製令報表";
                foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                {
                    ws.CreateRow(j + 1);

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ws.GetRow(j + 1).CreateCell(i).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[i].ToString());
                    }

                    j++;
                }
            }
            else if (tablename.Equals("TEMPds5"))
            {
                TABLENAME = "預計訂單完成報表";
                foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                {
                    ws.CreateRow(j + 1);

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ws.GetRow(j + 1).CreateCell(i).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[i].ToString());
                    }

                    j++;
                }
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

        public void SearchMATRIAL()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSqlMATERIAL();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    //dataGridView1.Columns.Clear();
                    ds2.Clear();

                    adapter.Fill(ds2, tablename);
                    sqlConn.Close();

                    if (ds2.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {

                        dataGridView2.DataSource = ds2.Tables[tablename];
                        dataGridView2.AutoResizeColumns();
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

        public StringBuilder SETsbSqlMATERIAL()
        {
            StringBuilder STR = new StringBuilder();


            if (!string.IsNullOrEmpty(comboBox2.Text.ToString()) && !comboBox2.Text.ToString().Equals("全部"))
            {

                STR.AppendFormat(@"   SELECT MD003 AS '品號',MB002 AS '品名',MD004 AS '單位',SUM(用量) AS '用量'");
                STR.AppendFormat(@"  ,( SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WITH(NOLOCK) WHERE LA009  IN ('20004','20006') AND  LA001=MD003) AS '現在庫存'");
                STR.AppendFormat(@"  ,(( SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WITH(NOLOCK) WHERE LA009  IN ('20004','20006') AND  LA001=MD003)-SUM(用量)) AS '可用量'");
                STR.AppendFormat(@"  ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WITH(NOLOCK) WHERE TD004=MD003 AND TD016='N' AND TD018='Y') AS '已採購量'");
                STR.AppendFormat(@"   FROM (");
                STR.AppendFormat(@"   SELECT MD003,[INVMB].MB002,MD004");
                STR.AppendFormat(@"   ,CONVERT(DECIMAL(18,4),(ISNULL([MOCMANULINE].NUM,0) +ISNULL([MOCMANULINE].BOX,0))/MC004*MD006/MD007) AS '用量'");
                STR.AppendFormat(@"   FROM [TK].dbo.[BOMMC],[TK].dbo.[BOMMD],[TK].dbo.[INVMB],[TKMOC].dbo.[MOCMANULINE]  ");
                STR.AppendFormat(@"   LEFT JOIN  [TKMOC].dbo.[MOCMANULINERESULT] ON [MOCMANULINE].ID=[MOCMANULINERESULT].SID  ");
                STR.AppendFormat(@"   WHERE [MOCMANULINE].MB001=MC001 AND MC001=MD001 AND MD003=[INVMB].MB001");
                STR.AppendFormat(@"   AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) >= '{0}'  AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) <= '{1}' ",dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"   AND [MOCMANULINE].MANU='{0}') AS TEMP",comboBox2.Text.ToString());
                STR.AppendFormat(@"   GROUP BY MD003,MB002,MD004");
                STR.AppendFormat(@"   ORDER BY MD003,MB002,MD004");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPdsMATERIAL1";
            }

            else if (comboBox2.Text.ToString().Equals("全部"))
            {


                STR.AppendFormat(@"   SELECT MD003 AS '品號',MB002 AS '品名',MD004 AS '單位',SUM(用量) AS '用量'");
                STR.AppendFormat(@"  ,( SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WITH(NOLOCK) WHERE LA009  IN ('20004','20006') AND  LA001=MD003) AS '現在庫存'");
                STR.AppendFormat(@"  ,(( SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WITH(NOLOCK) WHERE LA009  IN ('20004','20006') AND  LA001=MD003)-SUM(用量)) AS '可用量'");
                STR.AppendFormat(@"  ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WITH(NOLOCK) WHERE TD004=MD003 AND TD016='N' AND TD018='Y') AS '已採購量'");
                STR.AppendFormat(@"   FROM (");
                STR.AppendFormat(@"   SELECT MD003,[INVMB].MB002,MD004");
                STR.AppendFormat(@"   ,CONVERT(DECIMAL(18,4),(ISNULL([MOCMANULINE].NUM,0) +ISNULL([MOCMANULINE].BOX,0))/MC004*MD006/MD007) AS '用量'");
                STR.AppendFormat(@"   FROM [TK].dbo.[BOMMC],[TK].dbo.[BOMMD],[TK].dbo.[INVMB],[TKMOC].dbo.[MOCMANULINE]  ");
                STR.AppendFormat(@"   LEFT JOIN  [TKMOC].dbo.[MOCMANULINERESULT] ON [MOCMANULINE].ID=[MOCMANULINERESULT].SID  ");
                STR.AppendFormat(@"   WHERE [MOCMANULINE].MB001=MC001 AND MC001=MD001 AND MD003=[INVMB].MB001");
                STR.AppendFormat(@"   AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) >= '{0}'  AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) <= '{1}' ", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"   ) AS TEMP");
                STR.AppendFormat(@"   GROUP BY MD003,MB002,MD004");
                STR.AppendFormat(@"   ORDER BY MD003,MB002,MD004");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPdsMATERIAL2";
            }




            return STR;
        }

        public void ExcelExportMATERIAL()
        {
            SearchMATRIAL();
            string TABLENAME = "報表";
            int rows = 0;

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds2.Tables[tablename];

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
            if (tablename.Equals("TEMPdsMATERIAL1"))
            {
                TABLENAME = "預計原物料報表";
                foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                {
                    ws.CreateRow(j + 1);

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ws.GetRow(j + 1).CreateCell(i).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[i].ToString());
                    }

                    j++;
                }

            }
            else if (tablename.Equals("TEMPdsMATERIAL2"))
            {                
                TABLENAME = "預計原物料報表";
                foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                {
                    ws.CreateRow(j + 1);

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ws.GetRow(j + 1).CreateCell(i).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[i].ToString());
                    }

                    j++;
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    SOURCEID = row.Cells["ID"].Value.ToString();

                    SEARCHMOCMANULINECOP();
                }
                else
                {
                    SOURCEID = null;                 
                }
            }
        }

        public void SEARCHMOCMANULINECOP()
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

                sbSql.AppendFormat(@"  SELECT [MANU] AS '組別',[MOCMANULINECOP].[TC001] AS '訂單單別',[MOCMANULINECOP].[TC002] AS '訂單單號'");
                sbSql.AppendFormat(@"   ,[TC004] AS '客戶代號',[TC053] AS '客戶',[TC006] AS '業務',[MV002] AS '業務員'");
                sbSql.AppendFormat(@"   ,[SID] AS '來源',[ID]  ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINECOP] ");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[COPTC] ON [COPTC].[TC001]=[MOCMANULINECOP].[TC001] AND [COPTC].[TC002]=[MOCMANULINECOP].[TC002]");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[CMSMV] ON [MV001]=[TC006]");
                sbSql.AppendFormat(@"  WHERE [SID]='{0}'", SOURCEID);
                sbSql.AppendFormat(@"  ORDER BY [MANU],[MOCMANULINECOP].[TC001],[MOCMANULINECOP].[TC002]   ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");


                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView3.AutoResizeColumns();
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

        public void SEARCHMOCTG()
        { 
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql3();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    //dataGridView1.Columns.Clear();
                    ds.Clear();

                    adapter.Fill(ds, tablename);
                    sqlConn.Close();

                    if (ds.Tables[tablename].Rows.Count == 0)
                    {
                        dataGridView5.DataSource = null;
                    }
                    else
                    {

                        dataGridView5.DataSource = ds.Tables[tablename];
                        dataGridView5.AutoResizeColumns();
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

        public StringBuilder SETsbSql3()
        {
            StringBuilder STR = new StringBuilder();

            STR.AppendFormat(@"  SELECT TD013 AS '預交日',TC053 AS '客戶',TD004 AS '品號',TD005 AS '品名'");
            STR.AppendFormat(@"  ,ISNULL(CONVERT(DECIMAL(14,3),TD008*MD004/MD003),TD008) AS '下訂數量'");
            STR.AppendFormat(@"  ,MB004 AS '單位',TD008 AS '訂單量',TD010 AS '訂單單位'");
            STR.AppendFormat(@"  ,TC001 AS '訂單',TC002  AS '單號'");
            STR.AppendFormat(@"  ,(SELECT ISNULL(SUM(TG011),0) ");
            STR.AppendFormat(@"  FROM [TK].dbo.MOCTG,[TK].dbo.MOCTF ");
            STR.AppendFormat(@"  WHERE TG001=TF001 AND TG002=TF002");
            STR.AppendFormat(@"  AND TG009='1' ");
            STR.AppendFormat(@"  AND TF003<=TD013");
            STR.AppendFormat(@"  AND TG004=TD004");
            STR.AppendFormat(@"  AND TG014+TG015 IN  (");
            STR.AppendFormat(@"  SELECT[MOCMANULINERESULT].MOCTA001+[MOCMANULINERESULT].MOCTA002");
            STR.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINECOP],[TKMOC].[dbo].[MOCMANULINERESULT]");
            STR.AppendFormat(@"  WHERE [MOCMANULINECOP].[SID]=[MOCMANULINERESULT].[SID]");
            STR.AppendFormat(@"  AND [MOCMANULINECOP].TC001=[COPTC].TC001 AND [MOCMANULINECOP].TC002=[COPTC].TC002");
            STR.AppendFormat(@"  )) AS '實際入庫'");
            STR.AppendFormat(@"  FROM [TK].dbo.[COPTC],[TK].dbo.[COPTD]");
            STR.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD010");
            STR.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON MB001=TD004");
            STR.AppendFormat(@"  WHERE   COPTC.TC001=TD001 AND COPTC.TC002=TD002");
            STR.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            STR.AppendFormat(@"  AND TD008>0");
            STR.AppendFormat(@"  AND TD004 LIKE '4%'");
            STR.AppendFormat(@"  AND TD021='Y'");
            STR.AppendFormat(@"  ORDER BY TD013,COPTC.TC001,TD004,TD005");
            STR.AppendFormat(@"  ");
            STR.AppendFormat(@"  ");

            tablename = "TEMPds5";

            return STR;

        }
        public void SETCALENDAR()
        {
            string EVENT;
            DateTime dtEVENT;
            var ce2 = new CustomEvent();


            calendar1.RemoveAllEvents();
            calendar1.CalendarDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
            calendar1.CalendarView = CalendarViews.Month;
            calendar1.AllowEditingEvents = true;




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



                sbSql.AppendFormat(@"  SELECT [EVENTDATE],[MOCLINE],[EVENT]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[CALENDAR]");
                sbSql.AppendFormat(@"  WHERE [EVENTDATE]>='{0}'", DateTime.Now.ToString("yyyy") + "0101");
                sbSql.AppendFormat(@"  ORDER BY [EVENTDATE]");
                sbSql.AppendFormat(@"  ");

                adapterCALENDAR = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderCALENDAR = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                dsCALENDAR.Clear();
                adapterCALENDAR.Fill(dsCALENDAR, "TEMPdsCALENDAR");
                sqlConn.Close();


                if (dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows.Count >= 1)
                    {
                        foreach (DataRow od in dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows)
                        {
                            EVENT = od["MOCLINE"].ToString() + "-" + od["EVENT"].ToString();
                            dtEVENT = Convert.ToDateTime(od["EVENTDATE"].ToString());

                            ce2 = new CustomEvent
                            {
                                IgnoreTimeComponent = false,
                                EventText = EVENT,
                                Date = new DateTime(dtEVENT.Year, dtEVENT.Month, dtEVENT.Day),
                                EventLengthInHours = 2f,
                                RecurringFrequency = RecurringFrequencies.None,
                                EventFont = new Font("Verdana", 12, FontStyle.Regular),
                                Enabled = true,
                                EventColor = Color.FromArgb(120, 255, 120),
                                EventTextColor = Color.Black,
                                ThisDayForwardOnly = true
                            };

                            calendar1.AddEvent(ce2);
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

        public void SEARCHCOPTD()
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

                if (comboBox11.Text.Equals("未完成"))
                {
                    sbSqlQuery.AppendFormat(@" AND TD008-TD009>0 ");
                }
                else if (comboBox11.Text.Equals("已完成"))
                {
                    sbSqlQuery.AppendFormat(@" AND TD008-TD009=0 ");
                }
                else if (comboBox11.Text.Equals("全部"))
                {
                    sbSqlQuery.AppendFormat(@"  ");
                }



                sbSql.AppendFormat(@"  SELECT TD013 AS '預交日',TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD008 AS '訂單數',TD009 AS '已交數',TD010 AS '單位',TC053 AS '客戶'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTD,[TK].dbo.COPTC");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD001='A223'");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker12.Value.ToString("yyyyMMdd"), dateTimePicker13.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TD004 LIKE '401%'");
                sbSql.AppendFormat(@"  {0}", sbSqlQuery.ToString());
                sbSql.AppendFormat(@"  ORDER BY TD013,TD004");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter22 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder22 = new SqlCommandBuilder(adapter22);
                sqlConn.Open();
                ds22.Clear();
                adapter22.Fill(ds22, "TEMPds22");
                sqlConn.Close();


                if (ds22.Tables["TEMPds22"].Rows.Count == 0)
                {
                    dataGridView15.DataSource = null;
                }
                else
                {
                    if (ds22.Tables["TEMPds22"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView15.DataSource = ds22.Tables["TEMPds22"];
                        dataGridView15.AutoResizeColumns();
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
        private void dataGridView15_SelectionChanged(object sender, EventArgs e)
        {
            textBox49.Text = null;
            textBox50.Text = null;
            textBox51.Text = null;

            if (dataGridView15.CurrentRow != null)
            {
                int rowindex = dataGridView15.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView15.Rows[rowindex];
                    textBox49.Text = row.Cells["單別"].Value.ToString();
                    textBox50.Text = row.Cells["單號"].Value.ToString();
                    textBox51.Text = row.Cells["序號"].Value.ToString();
                }
                else
                {
                    textBox49.Text = null;
                    textBox50.Text = null;
                    textBox51.Text = null;
                }
            }
        }

        public void SETPATH()
        {
            DATES = DateTime.Now.ToString("yyyyMMddHHmmss");
            strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            pathFile = @""+strDesktopPath.ToString() + @"\"+"行事曆-預排" + DATES.ToString()+ comboBox4.Text.ToString();


            DeleteDir(pathFile + ".xlsx");
        }


        public void SETFILE()
        {


            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名 
            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();
            // 讓Excel文件可見
            //excelApp.Visible = true;
            // 停用警告訊息
            excelApp.DisplayAlerts = false;
            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);
            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];
            // 設定活頁簿焦點
            wBook.Activate();

            if (!File.Exists(pathFile + ".xlsx"))
            {
                wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }



            //關閉Excel
            excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            Console.Read();


            SEARCH();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}

        }

        public void CLEAREXCEL()
        {
            System.Diagnostics.Process[] p = System.Diagnostics.Process.GetProcesses();
            for (int i = 0; i < p.Length; i++)
            {
                if (p[i].ToString().IndexOf("EXCEL") > 0)
                    p[i].Kill();
            }
        }

        public void SEARCH()
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

                if(comboBox4.Text.Equals("包裝線"))
                {
                    sbSql.AppendFormat(@"  SELECT  CONVERT(NVARCHAR,[MANUDATE],112)+' ' +[MANU] AS MANUDATE,INVMB.[MB002],CONVERT(NVARCHAR,CONVERT(INT,ROUND([BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[PACKAGE]))+MB004 AS ' PACKAGE'  ");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE INVMB.MB001=MOCMANULINE.MB001");                    
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", sdt.ToString("yyyyMMdd"), edt.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MANU]='{0}'", comboBox4.Text);
                    sbSql.AppendFormat(@"  ORDER BY [MANU],[MANUDATE],MOCMANULINE.[MB001]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else
                {
                    sbSql.AppendFormat(@"  SELECT  CONVERT(NVARCHAR,[MANUDATE],112)+' ' +[MANU] AS MANUDATE,INVMB.[MB002],CONVERT(NVARCHAR,CONVERT(INT,ROUND([BAR],0)))+' 桶 '+CONVERT(NVARCHAR,CONVERT(INT,[NUM]))+MB004 AS ' PACKAGE'  ");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE INVMB.MB001=MOCMANULINE.MB001");                    
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", sdt.ToString("yyyyMMdd"), edt.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MANU]='{0}'", comboBox4.Text);
                    sbSql.AppendFormat(@"  ORDER BY [MANU],[MANUDATE],MOCMANULINE.[MB001]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }

              

                adapter3= new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = ds3.Tables["ds3"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    ds3.Tables["ds3"].Rows.Add(row);

                   // ExportDataSetToExcel(ds3, pathFile);
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(ds3, pathFile);
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

        public void ExportDataSetToExcel(DataSet ds, string TopathFile)
        {
            SETDATE();

            int days =Convert.ToInt32( sdt.AddDays(-sdt.Day + 1).DayOfWeek.ToString("d"));
            //MessageBox.Show(days.ToString());
            int MONTHDAYS= DateTime.DaysInMonth(sdt.Year, sdt.Month);

            int EXCELX = 2;
            int EXCELY = 0;

            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(TopathFile);
            Excel.Range wRange;
            Excel.Range wRangepathFile;
            Excel.Range wRangepathFilePURTA;

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        //if (table.Rows[j].ItemArray[0].ToString().Substring(6,2).Equals("01"))
                        if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 1)
                        {
                            message[0] = message[0] + table.Rows[j].ItemArray[k].ToString();
                            message[0] = message[0] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 2)
                        {
                            message[1] = message[1] + table.Rows[j].ItemArray[k].ToString();
                            message[1] = message[1] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 3)
                        {
                            message[2] = message[2] + table.Rows[j].ItemArray[k].ToString();
                            message[2] = message[2] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 4)
                        {
                            message[3] = message[3] + table.Rows[j].ItemArray[k].ToString();
                            message[3] = message[3] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 5)
                        {
                            message[4] = message[4] + table.Rows[j].ItemArray[k].ToString();
                            message[4] = message[4] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 6)
                        {
                            message[5] = message[5] + table.Rows[j].ItemArray[k].ToString();
                            message[5] = message[5] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 7)
                        {
                            message[6] = message[6] + table.Rows[j].ItemArray[k].ToString();
                            message[6] = message[6] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 8)
                        {
                            message[7] = message[7] + table.Rows[j].ItemArray[k].ToString();
                            message[7] = message[7] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 9)
                        {
                            message[8] = message[8] + table.Rows[j].ItemArray[k].ToString();
                            message[8] = message[8] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 10)
                        {
                            message[9] = message[9] + table.Rows[j].ItemArray[k].ToString();
                            message[9] = message[9] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 11)
                        {
                            message[10] = message[10] + table.Rows[j].ItemArray[k].ToString();
                            message[10] = message[10] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 12)
                        {
                            message[11] = message[11] + table.Rows[j].ItemArray[k].ToString();
                            message[11] = message[11] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 13)
                        {
                            message[12] = message[12] + table.Rows[j].ItemArray[k].ToString();
                            message[12] = message[12] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 14)
                        {
                            message[13] = message[13] + table.Rows[j].ItemArray[k].ToString();
                            message[13] = message[13] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 15)
                        {
                            message[14] = message[14] + table.Rows[j].ItemArray[k].ToString();
                            message[14] = message[14] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 16)
                        {
                            message[15] = message[15] + table.Rows[j].ItemArray[k].ToString();
                            message[15] = message[15] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 17)
                        {
                            message[16] = message[16] + table.Rows[j].ItemArray[k].ToString();
                            message[16] = message[16] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 18)
                        {
                            message[17] = message[17] + table.Rows[j].ItemArray[k].ToString();
                            message[17] = message[17] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 19)
                        {
                            message[18] = message[18] + table.Rows[j].ItemArray[k].ToString();
                            message[18] = message[18] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 20)
                        {
                            message[19] = message[19] + table.Rows[j].ItemArray[k].ToString();
                            message[19] = message[19] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 21)
                        {
                            message[20] = message[20] + table.Rows[j].ItemArray[k].ToString();
                            message[20] = message[20] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 22)
                        {
                            message[21] = message[21] + table.Rows[j].ItemArray[k].ToString();
                            message[21] = message[21] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 23)
                        {
                            message[22] = message[22] + table.Rows[j].ItemArray[k].ToString();
                            message[22] = message[22] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 24)
                        {
                            message[23] = message[23] + table.Rows[j].ItemArray[k].ToString();
                            message[23] = message[23] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 25)
                        {
                            message[24] = message[24] + table.Rows[j].ItemArray[k].ToString();
                            message[24] = message[24] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 26)
                        {
                            message[25] = message[25] + table.Rows[j].ItemArray[k].ToString();
                            message[25] = message[25] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 27)
                        {
                            message[26] = message[26] + table.Rows[j].ItemArray[k].ToString();
                            message[26] = message[26] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 28)
                        {
                            message[27] = message[27] + table.Rows[j].ItemArray[k].ToString();
                            message[27] = message[27] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 29)
                        {
                            message[28] = message[28] + table.Rows[j].ItemArray[k].ToString();
                            message[28] = message[28] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 30)
                        {
                            message[29] = message[29] + table.Rows[j].ItemArray[k].ToString();
                            message[29] = message[29] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 31)
                        {
                            message[30] = message[30] + table.Rows[j].ItemArray[k].ToString();
                            message[30] = message[30] + '\n';
                        }
                    }
                    //message = message + '\n';
                }

                excelWorkSheet.Cells[1, 1] = "星期日";
                excelWorkSheet.Cells[1, 2] = "星期一";
                excelWorkSheet.Cells[1, 3] = "星期二";
                excelWorkSheet.Cells[1, 4] = "星期三";
                excelWorkSheet.Cells[1, 5] = "星期四";
                excelWorkSheet.Cells[1, 6] = "星期五";
                excelWorkSheet.Cells[1, 7] = "星期六";

                //置中
                string RangeCenter = "A1:G1";//設定範圍
                excelWorkSheet.get_Range(RangeCenter).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                for (int i = 1; i <= MONTHDAYS;i++)
                {
                    
                    EXCELX = 2 + Convert.ToInt32(Math.Truncate(Convert.ToDouble((i+days-1) / 7)));
                    EXCELY = (days + i) % 7;
                    if(EXCELY==0)
                    {
                        EXCELY = 7;                        
                    }

                    //excelWorkSheet.Cells[EXCELX, EXCELY] = i;

                    excelWorkSheet.Cells[EXCELX, EXCELY] = message[i-1].ToString();

                    //if (!string.IsNullOrEmpty(message[i-1].ToString()))
                    //{
                    //    excelWorkSheet.Cells[EXCELX, EXCELY] = message[i - 1].ToString();
                    //}

                }
                //excelWorkSheet.Cells[1, 1] = dateTimePicker9.Value.ToString("yyyy/MM/") + "01";
                //excelWorkSheet.Cells[2, days+1] = message1;
                //message1 = null;
                

                //靠左
                string RangeLeft = "A2:G6";//設定範圍
                excelWorkSheet.get_Range(RangeLeft).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                //設定為按照內容自動調整欄寬
                //excelWorkSheet.get_Range(RangeLeft).Columns.AutoFit();
                excelWorkSheet.get_Range(RangeLeft).ColumnWidth = 30;
                //excelWorkSheet.Columns.AutoFit();

                // 給儲存格加邊框
                excelWorkSheet.get_Range(RangeLeft).Borders.LineStyle = Excel.XlBorderWeight.xlHairline;
                //excelWorkSheet.get_Range(RangeLeft).Borders.LineStyle = Excel.XlBorderWeight.xlMedium;
            }



            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();

        }

        public void SETPATH2()
        {
            DATES = DateTime.Now.ToString("yyyyMMddHHmmss");
            strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            pathFile2 = @"" + strDesktopPath.ToString() + @"\" + "行事曆-製令" + DATES.ToString() + comboBox5.Text.ToString();


            DeleteDir(pathFile2 + ".xlsx");
        }

        public void SETPATH3()
        {
            DATES = DateTime.Now.ToString("yyyyMMddHHmmss");
            strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            pathFile2 = @"" + strDesktopPath.ToString() + @"\" + "行事曆-製令" + DATES.ToString() + comboBox6.Text.ToString();


            DeleteDir(pathFile2 + ".xlsx");
        }

        public void SETPATH4()
        {
            DATES = DateTime.Now.ToString("yyyyMMddHHmmss");
            strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            pathFile4 = @"" + strDesktopPath.ToString() + @"\" + "行事曆-製令-工時" + DATES.ToString() + comboBox6.Text.ToString();


            DeleteDir(pathFile4 + ".xlsx");
        }
        public void SETFILE2()
        {


            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名 
            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();
            // 讓Excel文件可見
            //excelApp.Visible = true;
            // 停用警告訊息
            excelApp.DisplayAlerts = false;
            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);
            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];
            // 設定活頁簿焦點
            wBook.Activate();

            if (!File.Exists(pathFile2 + ".xlsx"))
            {
                wBook.SaveAs(pathFile2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }



            //關閉Excel
            excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            Console.Read();


            SEARCH2();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}

        }

        public void SETFILE3()
        {


            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名 
            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();
            // 讓Excel文件可見
            //excelApp.Visible = true;
            // 停用警告訊息
            excelApp.DisplayAlerts = false;
            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);
            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];
            // 設定活頁簿焦點
            wBook.Activate();

            if (!File.Exists(pathFile2 + ".xlsx"))
            {
                wBook.SaveAs(pathFile2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }



            //關閉Excel
            excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            Console.Read();


            SEARCH3();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}

        }


        public void SETFILE4()
        {


            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名 
            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();
            // 讓Excel文件可見
            //excelApp.Visible = true;
            // 停用警告訊息
            excelApp.DisplayAlerts = false;
            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);
            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];
            // 設定活頁簿焦點
            wBook.Activate();

            if (!File.Exists(pathFile4 + ".xlsx"))
            {
                wBook.SaveAs(pathFile4, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }



            //關閉Excel
            excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            Console.Read();


            SEARCH4();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}

        }
        public void CLEAREXCEL2()
        {
            System.Diagnostics.Process[] p = System.Diagnostics.Process.GetProcesses();
            for (int i = 0; i < p.Length; i++)
            {
                if (p[i].ToString().IndexOf("EXCEL") > 0)
                    p[i].Kill();
            }
        }

        public void SEARCH2()
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

                sbSql.AppendFormat(@"  SELECT  CONVERT(NVARCHAR,TA009,112)+' ' +MD002 AS MANUDATE,INVMB.[MB002],CONVERT(NVARCHAR,CONVERT(INT,ROUND(TA015,0)))++' '+TA007 AS ' PACKAGE'    ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA,[TK].dbo.INVMB,[TK].dbo.CMSMD ");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=MD001");
                sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,TA009,112) >='{0}' AND CONVERT(NVARCHAR,TA009,112) <='{1}'",sdt2.ToString("yyyyMMdd"), edt2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'  ",comboBox5.Text.ToString());
                sbSql.AppendFormat(@" ORDER BY MD002,[MANUDATE],MB001  ");
                sbSql.AppendFormat(@"  ");

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "ds5");
                sqlConn.Close();


                if (ds5.Tables["ds5"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = ds5.Tables["ds5"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    ds5.Tables["ds5"].Rows.Add(row);

                    //ExportDataSetToExcel2(ds5, pathFile2);
                }
                else
                {
                    if (ds5.Tables["ds5"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel2(ds5, pathFile2);
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
        public void SEARCH3()
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

                StringBuilder SB = new StringBuilder();

                if (comboBox6.Text.Equals("包裝線"))
                {
                    sbSql.AppendFormat(@"  SELECT  CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)+' '+[MOCMANULINE].[MANU] AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004 AS ' PACKAGE'");
                    sbSql.AppendFormat(@"  ,'---' AS '-'");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                    sbSql.AppendFormat(@"  WHERE INVMB.MB001=MOCMANULINE.MB001    ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='包裝線'");
                    sbSql.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU],[MANUDATE]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox6.Text.Equals("製二線"))
                {
                    sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)+' '+[MOCMANULINE].[MANU] AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS ' MB002', CONVERT(NVARCHAR,CONVERT(INT,ROUND([NUM],0)))++' KG' AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,'---' AS '-' ");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB ");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004            ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='製二線'");
                    sbSql.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox6.Text.Equals("製一線"))
                {
                    sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)+' '+[MOCMANULINE].[MANU] AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS ' MB002', CONVERT(NVARCHAR,CONVERT(INT,ROUND([NUM],0)))+' '+INVMB.MB004 AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,'---' AS '-' ");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB ");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004            ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='製一線'");
                    sbSql.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox6.Text.Equals("手工線"))
                {
                    sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)+' '+[MOCMANULINE].[MANU] AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS ' MB002', CONVERT(NVARCHAR,CONVERT(INT,ROUND([NUM],0)))++' KG' AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,'---' AS '-' ");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB ");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004            ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='手工線'");
                    sbSql.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }

                sbSql.AppendFormat(@"  ");

                adapter6 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder6 = new SqlCommandBuilder(adapter6);
                sqlConn.Open();
                ds6.Clear();
                adapter6.Fill(ds6, "ds6");
                sqlConn.Close();


                if (ds6.Tables["ds6"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = ds6.Tables["ds6"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    ds6.Tables["ds6"].Rows.Add(row);

                    //ExportDataSetToExcel2(ds5, pathFile2);
                }
                else
                {
                    if (ds6.Tables["ds6"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel2(ds6, pathFile2);
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

        public void SEARCH4()
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

                StringBuilder SB = new StringBuilder();

                if (comboBox6.Text.Equals("包裝線"))
                {
                    sbSql.AppendFormat(@"  SELECT MANUDATE,MANU,PACKAGE,HRS,REMARK");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE,'-總工時- ' AS 'PACKAGE',SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))) AS HRS,'---' AS 'REMARK'");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                    sbSql.AppendFormat(@"  WHERE INVMB.MB001=MOCMANULINE.MB001    ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='包裝線'");
                    sbSql.AppendFormat(@"  GROUP BY [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) ");
                    sbSql.AppendFormat(@"  UNION ALL");
                    sbSql.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004 AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)) AS HRS");
                    sbSql.AppendFormat(@"  ,'---' AS 'REMARK'");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                    sbSql.AppendFormat(@"  WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='包裝線'");
                    sbSql.AppendFormat(@"  ) AS TEMP1");
                    sbSql.AppendFormat(@"  ORDER BY  [MANU],CONVERT(NVARCHAR,[MANUDATE],112),PACKAGE");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox6.Text.Equals("製二線"))
                {
                    sbSql.AppendFormat(@"  SELECT MANUDATE,MANU,PACKAGE,HRS,REMARK");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE,'-總工時- ' AS 'PACKAGE',SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),2),0))) AS HRS,'---' AS 'REMARK'");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.COPTD,[TK].dbo.INVMB,[TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=[MOCMANULINE].MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='製二線'");
                    sbSql.AppendFormat(@"  GROUP BY [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) ");
                    sbSql.AppendFormat(@"  UNION ALL");
                    sbSql.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),2),0) AS HRS");
                    sbSql.AppendFormat(@"  ,'-'+[MOCMANULINE].MB001");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.COPTD,[TK].dbo.INVMB,[TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=[MOCMANULINE].MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='製二線'");
                    sbSql.AppendFormat(@"  ) AS TEMP1");
                    sbSql.AppendFormat(@"  ORDER BY  [MANU],CONVERT(NVARCHAR,[MANUDATE],112),PACKAGE");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox6.Text.Equals("製一線"))
                {

                    sbSql.AppendFormat(@"  SELECT MANUDATE,MANU,PACKAGE,HRS,REMARK");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE,'-總工時- ' AS 'PACKAGE',SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),2),0))) AS HRS,'---' AS 'REMARK'");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.COPTD,[TK].dbo.INVMB,[TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=[MOCMANULINE].MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='製一線'");
                    sbSql.AppendFormat(@"  GROUP BY [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) ");
                    sbSql.AppendFormat(@"  UNION ALL");
                    sbSql.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),2),0) AS HRS");
                    sbSql.AppendFormat(@"  ,'-'+[MOCMANULINE].MB001");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.COPTD,[TK].dbo.INVMB,[TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=[MOCMANULINE].MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='製一線'");
                    sbSql.AppendFormat(@"  ) AS TEMP1");
                    sbSql.AppendFormat(@"  ORDER BY  [MANU],CONVERT(NVARCHAR,[MANUDATE],112),PACKAGE");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox6.Text.Equals("手工線"))
                {

                    sbSql.AppendFormat(@"  SELECT MANUDATE,MANU,PACKAGE,HRS,REMARK");
                    sbSql.AppendFormat(@"  FROM (");
                    sbSql.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE,'-總工時- ' AS 'PACKAGE',SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),2),0))) AS HRS,'---' AS 'REMARK'");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.COPTD,[TK].dbo.INVMB,[TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=[MOCMANULINE].MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='手工線'");
                    sbSql.AppendFormat(@"  GROUP BY [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) ");
                    sbSql.AppendFormat(@"  UNION ALL");
                    sbSql.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),2),0) AS HRS");
                    sbSql.AppendFormat(@"  ,'-'+[MOCMANULINE].MB001");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.COPTD,[TK].dbo.INVMB,[TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=[MOCMANULINE].MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='手工線'");
                    sbSql.AppendFormat(@"  ) AS TEMP1");
                    sbSql.AppendFormat(@"  ORDER BY  [MANU],CONVERT(NVARCHAR,[MANUDATE],112),PACKAGE");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }

                sbSql.AppendFormat(@"  ");

                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "ds7");
                sqlConn.Close();


                if (ds7.Tables["ds7"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = ds7.Tables["ds7"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    ds7.Tables["ds7"].Rows.Add(row);

                    //ExportDataSetToExcel2(ds5, pathFile2);
                }
                else
                {
                    if (ds7.Tables["ds7"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel4(ds7, pathFile4,message4);
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

        public void ExportDataSetToExcel2(DataSet ds, string TopathFile)
        {
            SETDATE4();

            int days = Convert.ToInt32(sdt4.AddDays(-sdt4.Day + 1).DayOfWeek.ToString("d"));
            //MessageBox.Show(days.ToString());
            int MONTHDAYS = DateTime.DaysInMonth(sdt4.Year, sdt4.Month);

            int EXCELX = 2;
            int EXCELY = 0;

            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(TopathFile);
            Excel.Range wRange;
            Excel.Range wRangepathFile;
            Excel.Range wRangepathFilePURTA;

           

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        //if (table.Rows[j].ItemArray[0].ToString().Substring(6,2).Equals("01"))
                        if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 1)
                        {
                            message3[0] = message3[0] + table.Rows[j].ItemArray[k].ToString();
                            message3[0] = message3[0] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 2)
                        {
                            message3[1] = message3[1] + table.Rows[j].ItemArray[k].ToString();
                            message3[1] = message3[1] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 3)
                        {
                            message3[2] = message3[2] + table.Rows[j].ItemArray[k].ToString();
                            message3[2] = message3[2] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 4)
                        {
                            message3[3] = message3[3] + table.Rows[j].ItemArray[k].ToString();
                            message3[3] = message3[3] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 5)
                        {
                            message3[4] = message3[4] + table.Rows[j].ItemArray[k].ToString();
                            message3[4] = message3[4] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 6)
                        {
                            message3[5] = message3[5] + table.Rows[j].ItemArray[k].ToString();
                            message3[5] = message3[5] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 7)
                        {
                            message3[6] = message3[6] + table.Rows[j].ItemArray[k].ToString();
                            message3[6] = message3[6] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 8)
                        {
                            message3[7] = message3[7] + table.Rows[j].ItemArray[k].ToString();
                            message3[7] = message3[7] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 9)
                        {
                            message3[8] = message3[8] + table.Rows[j].ItemArray[k].ToString();
                            message3[8] = message3[8] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 10)
                        {
                            message3[9] = message3[9] + table.Rows[j].ItemArray[k].ToString();
                            message3[9] = message3[9] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 11)
                        {
                            message3[10] = message3[10] + table.Rows[j].ItemArray[k].ToString();
                            message3[10] = message3[10] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 12)
                        {
                            message3[11] = message3[11] + table.Rows[j].ItemArray[k].ToString();
                            message3[11] = message3[11] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 13)
                        {
                            message3[12] = message3[12] + table.Rows[j].ItemArray[k].ToString();
                            message3[12] = message3[12] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 14)
                        {
                            message3[13] = message3[13] + table.Rows[j].ItemArray[k].ToString();
                            message3[13] = message3[13] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 15)
                        {
                            message3[14] = message3[14] + table.Rows[j].ItemArray[k].ToString();
                            message3[14] = message3[14] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 16)
                        {
                            message3[15] = message3[15] + table.Rows[j].ItemArray[k].ToString();
                            message3[15] = message3[15] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 17)
                        {
                            message3[16] = message3[16] + table.Rows[j].ItemArray[k].ToString();
                            message3[16] = message3[16] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 18)
                        {
                            message3[17] = message3[17] + table.Rows[j].ItemArray[k].ToString();
                            message3[17] = message3[17] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 19)
                        {
                            message3[18] = message3[18] + table.Rows[j].ItemArray[k].ToString();
                            message3[18] = message3[18] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 20)
                        {
                            message3[19] = message3[19] + table.Rows[j].ItemArray[k].ToString();
                            message3[19] = message3[19] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 21)
                        {
                            message3[20] = message3[20] + table.Rows[j].ItemArray[k].ToString();
                            message3[20] = message3[20] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 22)
                        {
                            message3[21] = message3[21] + table.Rows[j].ItemArray[k].ToString();
                            message3[21] = message3[21] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 23)
                        {
                            message3[22] = message3[22] + table.Rows[j].ItemArray[k].ToString();
                            message3[22] = message3[22] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 24)
                        {
                            message3[23] = message3[23] + table.Rows[j].ItemArray[k].ToString();
                            message3[23] = message3[23] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 25)
                        {
                            message3[24] = message3[24] + table.Rows[j].ItemArray[k].ToString();
                            message3[24] = message3[24] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 26)
                        {
                            message3[25] = message3[25] + table.Rows[j].ItemArray[k].ToString();
                            message3[25] = message3[25] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 27)
                        {
                            message3[26] = message3[26] + table.Rows[j].ItemArray[k].ToString();
                            message3[26] = message3[26] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 28)
                        {
                            message3[27] = message3[27] + table.Rows[j].ItemArray[k].ToString();
                            message3[27] = message3[27] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 29)
                        {
                            message3[28] = message3[28] + table.Rows[j].ItemArray[k].ToString();
                            message3[28] = message3[28] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 30)
                        {
                            message3[29] = message3[29] + table.Rows[j].ItemArray[k].ToString();
                            message3[29] = message3[29] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 31)
                        {
                            message3[30] = message3[30] + table.Rows[j].ItemArray[k].ToString();
                            message3[30] = message3[30] + '\n';
                        }
                    }
                    //message = message + '\n';
                }

                excelWorkSheet.Cells[1, 1] = "星期日";
                excelWorkSheet.Cells[1, 2] = "星期一";
                excelWorkSheet.Cells[1, 3] = "星期二";
                excelWorkSheet.Cells[1, 4] = "星期三";
                excelWorkSheet.Cells[1, 5] = "星期四";
                excelWorkSheet.Cells[1, 6] = "星期五";
                excelWorkSheet.Cells[1, 7] = "星期六";

                //置中
                string RangeCenter = "A1:G1";//設定範圍
                excelWorkSheet.get_Range(RangeCenter).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                for (int i = 1; i <= MONTHDAYS; i++)
                {

                    EXCELX = 2 + Convert.ToInt32(Math.Truncate(Convert.ToDouble((i + days - 1) / 7)));
                    EXCELY = (days + i) % 7;
                    if (EXCELY == 0)
                    {
                        EXCELY = 7;
                    }

                    //excelWorkSheet.Cells[EXCELX, EXCELY] = i;

                    excelWorkSheet.Cells[EXCELX, EXCELY] = message3[i - 1].ToString();

                    //if (!string.IsNullOrEmpty(message[i-1].ToString()))
                    //{
                    //    excelWorkSheet.Cells[EXCELX, EXCELY] = message[i - 1].ToString();
                    //}

                }
                //excelWorkSheet.Cells[1, 1] = dateTimePicker9.Value.ToString("yyyy/MM/") + "01";
                //excelWorkSheet.Cells[2, days+1] = message1;
                //message1 = null;


                //靠左
                string RangeLeft = "A2:G6";//設定範圍
                excelWorkSheet.get_Range(RangeLeft).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                //設定為按照內容自動調整欄寬
                //excelWorkSheet.get_Range(RangeLeft).Columns.AutoFit();
                excelWorkSheet.get_Range(RangeLeft).ColumnWidth = 30;
                //excelWorkSheet.Columns.AutoFit();

                // 給儲存格加邊框
                excelWorkSheet.get_Range(RangeLeft).Borders.LineStyle = Excel.XlBorderWeight.xlHairline;
                //excelWorkSheet.get_Range(RangeLeft).Borders.LineStyle = Excel.XlBorderWeight.xlMedium;
            }



            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();

        }

        public void ExportDataSetToExcel4(DataSet ds, string TopathFile, string[] message4)
        {
            SETDATE4();

            int days = Convert.ToInt32(sdt4.AddDays(-sdt4.Day + 1).DayOfWeek.ToString("d"));
            //MessageBox.Show(days.ToString());
            int MONTHDAYS = DateTime.DaysInMonth(sdt4.Year, sdt4.Month);

            int EXCELX = 2;
            int EXCELY = 0;

            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(TopathFile);
            Excel.Range wRange;
            Excel.Range wRangepathFile;
            Excel.Range wRangepathFilePURTA;



            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        //if (table.Rows[j].ItemArray[0].ToString().Substring(6,2).Equals("01"))
                        if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 1)
                        {
                            message4[0] = message4[0] + table.Rows[j].ItemArray[k].ToString();
                            message4[0] = message4[0] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 2)
                        {
                            message4[1] = message4[1] + table.Rows[j].ItemArray[k].ToString();
                            message4[1] = message4[1] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 3)
                        {
                            message4[2] = message4[2] + table.Rows[j].ItemArray[k].ToString();
                            message4[2] = message4[2] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 4)
                        {
                            message4[3] = message4[3] + table.Rows[j].ItemArray[k].ToString();
                            message4[3] = message4[3] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 5)
                        {
                            message4[4] = message4[4] + table.Rows[j].ItemArray[k].ToString();
                            message4[4] = message4[4] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 6)
                        {
                            message4[5] = message4[5] + table.Rows[j].ItemArray[k].ToString();
                            message4[5] = message4[5] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 7)
                        {
                            message4[6] = message4[6] + table.Rows[j].ItemArray[k].ToString();
                            message4[6] = message4[6] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 8)
                        {
                            message4[7] = message4[7] + table.Rows[j].ItemArray[k].ToString();
                            message4[7] = message4[7] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 9)
                        {
                            message4[8] = message4[8] + table.Rows[j].ItemArray[k].ToString();
                            message4[8] = message4[8] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 10)
                        {
                            message4[9] = message4[9] + table.Rows[j].ItemArray[k].ToString();
                            message4[9] = message4[9] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 11)
                        {
                            message4[10] = message4[10] + table.Rows[j].ItemArray[k].ToString();
                            message4[10] = message4[10] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 12)
                        {
                            message4[11] = message4[11] + table.Rows[j].ItemArray[k].ToString();
                            message4[11] = message4[11] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 13)
                        {
                            message4[12] = message4[12] + table.Rows[j].ItemArray[k].ToString();
                            message4[12] = message4[12] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 14)
                        {
                            message4[13] = message4[13] + table.Rows[j].ItemArray[k].ToString();
                            message4[13] = message4[13] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 15)
                        {
                            message4[14] = message4[14] + table.Rows[j].ItemArray[k].ToString();
                            message4[14] = message4[14] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 16)
                        {
                            message4[15] = message4[15] + table.Rows[j].ItemArray[k].ToString();
                            message4[15] = message4[15] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 17)
                        {
                            message4[16] = message4[16] + table.Rows[j].ItemArray[k].ToString();
                            message4[16] = message4[16] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 18)
                        {
                            message4[17] = message4[17] + table.Rows[j].ItemArray[k].ToString();
                            message4[17] = message4[17] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 19)
                        {
                            message4[18] = message4[18] + table.Rows[j].ItemArray[k].ToString();
                            message4[18] = message4[18] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 20)
                        {
                            message4[19] = message4[19] + table.Rows[j].ItemArray[k].ToString();
                            message4[19] = message4[19] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 21)
                        {
                            message4[20] = message4[20] + table.Rows[j].ItemArray[k].ToString();
                            message4[20] = message4[20] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 22)
                        {
                            message4[21] = message4[21] + table.Rows[j].ItemArray[k].ToString();
                            message4[21] = message4[21] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 23)
                        {
                            message4[22] = message4[22] + table.Rows[j].ItemArray[k].ToString();
                            message4[22] = message4[22] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 24)
                        {
                            message4[23] = message4[23] + table.Rows[j].ItemArray[k].ToString();
                            message4[23] = message4[23] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 25)
                        {
                            message4[24] = message4[24] + table.Rows[j].ItemArray[k].ToString();
                            message4[24] = message4[24] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 26)
                        {
                            message4[25] = message4[25] + table.Rows[j].ItemArray[k].ToString();
                            message4[25] = message4[25] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 27)
                        {
                            message4[26] = message4[26] + table.Rows[j].ItemArray[k].ToString();
                            message4[26] = message4[26] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 28)
                        {
                            message4[27] = message4[27] + table.Rows[j].ItemArray[k].ToString();
                            message4[27] = message4[27] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 29)
                        {
                            message4[28] = message4[28] + table.Rows[j].ItemArray[k].ToString();
                            message4[28] = message4[28] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 30)
                        {
                            message4[29] = message4[29] + table.Rows[j].ItemArray[k].ToString();
                            message4[29] = message4[29] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 31)
                        {
                            message4[30] = message4[30] + table.Rows[j].ItemArray[k].ToString();
                            message4[30] = message4[30] + '\n';
                        }
                    }
                    //message = message + '\n';
                }

                excelWorkSheet.Cells[1, 1] = "星期日";
                excelWorkSheet.Cells[1, 2] = "星期一";
                excelWorkSheet.Cells[1, 3] = "星期二";
                excelWorkSheet.Cells[1, 4] = "星期三";
                excelWorkSheet.Cells[1, 5] = "星期四";
                excelWorkSheet.Cells[1, 6] = "星期五";
                excelWorkSheet.Cells[1, 7] = "星期六";

                //置中
                string RangeCenter = "A1:G1";//設定範圍
                excelWorkSheet.get_Range(RangeCenter).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                for (int i = 1; i <= MONTHDAYS; i++)
                {

                    EXCELX = 2 + Convert.ToInt32(Math.Truncate(Convert.ToDouble((i + days - 1) / 7)));
                    EXCELY = (days + i) % 7;
                    if (EXCELY == 0)
                    {
                        EXCELY = 7;
                    }

                    //excelWorkSheet.Cells[EXCELX, EXCELY] = i;

                    excelWorkSheet.Cells[EXCELX, EXCELY] = message4[i - 1].ToString();

                    //if (!string.IsNullOrEmpty(message[i-1].ToString()))
                    //{
                    //    excelWorkSheet.Cells[EXCELX, EXCELY] = message[i - 1].ToString();
                    //}

                }
                //excelWorkSheet.Cells[1, 1] = dateTimePicker9.Value.ToString("yyyy/MM/") + "01";
                //excelWorkSheet.Cells[2, days+1] = message1;
                //message1 = null;


                //靠左
                string RangeLeft = "A2:G6";//設定範圍
                excelWorkSheet.get_Range(RangeLeft).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                //設定為按照內容自動調整欄寬
                //excelWorkSheet.get_Range(RangeLeft).Columns.AutoFit();
                excelWorkSheet.get_Range(RangeLeft).ColumnWidth = 30;
                //excelWorkSheet.Columns.AutoFit();

                // 給儲存格加邊框
                excelWorkSheet.get_Range(RangeLeft).Borders.LineStyle = Excel.XlBorderWeight.xlHairline;
                //excelWorkSheet.get_Range(RangeLeft).Borders.LineStyle = Excel.XlBorderWeight.xlMedium;
            }



            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();

        }


        public void DeleteDir(string aimPath)
        {
            try
            {
                File.Delete(aimPath);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public void RESET()
        {
            message = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
           
        }
        public void RESET2()
        {
            message2 = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };

        }

        public void RESET3()
        {
            message3 = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };

        }
        public void RESET4()
        {
            message4 = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };

        }
        private void dateTimePicker9_ValueChanged(object sender, EventArgs e)
        {
            SETDATE();
        }

        private void dateTimePicker10_ValueChanged(object sender, EventArgs e)
        {
            SETDATE2();
        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1();
            Report report1 = new Report();
            report1.Load(@"REPORT\預排訂單行事曆.frx");

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

            if (comboBox6.Text.Equals("包裝線"))
            {
                SB.AppendFormat(@"  SELECT  [PREINVMBMANU].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004 AS 'PACKAGE'");
                SB.AppendFormat(@"  ,ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),2),0) AS HRS");
                SB.AppendFormat(@"  ,INVMB.MB001");
                SB.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                SB.AppendFormat(@"  WHERE INVMB.MB001=MOCMANULINE.MB001      ");
                SB.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ",dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='包裝線'");
                SB.AppendFormat(@"  ORDER BY [PREINVMBMANU].[MANU],[MOCMANULINE].[MANU],[MANUDATE]  ");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"   ");
                SB.AppendFormat(@"  ");
            }
            else if (comboBox6.Text.Equals("製二線"))
            {
                SB.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS 'PACKAGE'");
                SB.AppendFormat(@"  ,ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),2),0) AS HRS");
                SB.AppendFormat(@"  ,INVMB.MB001");
                SB.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB");
                SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'");
                SB.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                SB.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                SB.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='製二線'");
                SB.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
            }
            else if (comboBox6.Text.Equals("製一線"))
            {
                SB.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS 'PACKAGE'");
                SB.AppendFormat(@"  ,ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),2),0) AS HRS");
                SB.AppendFormat(@"  ,INVMB.MB001");
                SB.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB");
                SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'");
                SB.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                SB.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                SB.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='製一線'");
                SB.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
            }
            else if (comboBox6.Text.Equals("手工線"))
            {
                SB.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS 'PACKAGE'");
                SB.AppendFormat(@"  ,ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),2),0) AS HRS");
                SB.AppendFormat(@"  ,INVMB.MB001");
                SB.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB");
                SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'");
                SB.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                SB.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                SB.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='手工線'");
                SB.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"   ");
            }


            return SB;

        }

        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2();
            Report report1 = new Report();
            report1.Load(@"REPORT\生產-檢查訂單是否有預排.frx");

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

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  SELECT TC053,TD013");
            SB.AppendFormat(@"  ,CONVERT(NVARCHAR,MOCMANULINE1.[MANUDATE],112)  AS '包裝線生產日'");
            SB.AppendFormat(@"  ,CONVERT(NVARCHAR,MOCMANULINE2.[MANUDATE],112)  AS '製一線生產日'");
            SB.AppendFormat(@"  ,CONVERT(NVARCHAR,MOCMANULINE3.[MANUDATE],112)  AS '製二線生產日'");
            SB.AppendFormat(@"  ,CONVERT(NVARCHAR,MOCMANULINE4.[MANUDATE],112)  AS '手工線生產日'");
            SB.AppendFormat(@"  ,TD001,TD002,TD003,TD004,TD005,TD006,TD008,TD009,TD024,TD025");
            SB.AppendFormat(@"  ,CASE WHEN MD002=TD010 THEN MD004*(TD008-TD009+TD024-TD025) ELSE (TD008-TD009+TD024-TD025) END AS 'NUM'");
            SB.AppendFormat(@"  ,MOCMANULINE1.[PACKAGE] '包裝線生產數'");
            SB.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD010");
            SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MOCMANULINE] MOCMANULINE1 ON MOCMANULINE1.[MANU]='包裝線' AND MOCMANULINE1.[COPTD001]=TD001 AND MOCMANULINE1.[COPTD002]=TD002 AND MOCMANULINE1.[COPTD003]=TD003");
            SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MOCMANULINE] MOCMANULINE2 ON MOCMANULINE2.[MANU]='製一線' AND MOCMANULINE2.[COPTD001]=TD001 AND MOCMANULINE2.[COPTD002]=TD002 AND MOCMANULINE2.[COPTD003]=TD003");
            SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MOCMANULINE] MOCMANULINE3 ON MOCMANULINE3.[MANU]='製二線' AND MOCMANULINE3.[COPTD001]=TD001 AND MOCMANULINE3.[COPTD002]=TD002 AND MOCMANULINE3.[COPTD003]=TD003");
            SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[MOCMANULINE] MOCMANULINE4 ON MOCMANULINE4.[MANU]='手工線' AND MOCMANULINE4.[COPTD001]=TD001 AND MOCMANULINE4.[COPTD002]=TD002 AND MOCMANULINE4.[COPTD003]=TD003");
            SB.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(@"  AND COPTD.UDF01='Y' AND TD016='N' AND TD021='Y'");
            SB.AppendFormat(@"  AND (TD004 LIKE '4%' OR TD004 LIKE '5%')");
            SB.AppendFormat(@"  AND TD013>='{0}' AND  TD013<='{1}'",dateTimePicker15.Value.ToString("yyyyMMdd"), dateTimePicker16.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@"  ORDER BY TC053,TD013,TD001,TD002,TD003");
            SB.AppendFormat(@"  ");
            SB.AppendFormat(@"  ");
            SB.AppendFormat(@"  ");


            return SB;

        }

        public void SETFASTREPORT3()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3();
            Report report1 = new Report();
            report1.Load(@"REPORT\生產-檢查訂單是否有製令.frx");

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

            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  SELECT TC053,TD013");
            SB.AppendFormat(@"  ,CONVERT(NVARCHAR,MOCTA1.TA009,112)+' '+MOCTA1.TA001+'-'+MOCTA1.TA002  AS '包裝線生產日'");
            SB.AppendFormat(@"  ,CONVERT(NVARCHAR,MOCTA2.TA009,112)+' '+MOCTA2.TA001+'-'+MOCTA2.TA002   AS '製一線生產日'");
            SB.AppendFormat(@"  ,CONVERT(NVARCHAR,MOCTA3.TA009,112)+' '+MOCTA3.TA001+'-'+MOCTA3.TA002   AS '製二線生產日'");
            SB.AppendFormat(@"  ,CONVERT(NVARCHAR,MOCTA4.TA009,112)+' '+MOCTA4.TA001+'-'+MOCTA4.TA002   AS '手工線生產日'");
            SB.AppendFormat(@"  ,TD001,TD002,TD003,TD004,TD005,TD006,TD008,TD009,TD024,TD025");
            SB.AppendFormat(@"  ,CASE WHEN MD002=TD010 THEN MD004*(TD008-TD009+TD024-TD025) ELSE (TD008-TD009+TD024-TD025) END AS 'NUM'");
            SB.AppendFormat(@"  ,MOCTA1.TA015 '包裝線生產數'");
            SB.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD010");
            SB.AppendFormat(@"  LEFT JOIN [TK].dbo.MOCTA MOCTA1 ON MOCTA1.TA021='09' AND MOCTA1.TA026=TD001 AND MOCTA1.TA027=TD002 AND MOCTA1.TA028=TD003");
            SB.AppendFormat(@"  LEFT JOIN [TK].dbo.MOCTA MOCTA2 ON MOCTA2.TA021='01' AND MOCTA2.TA026=TD001 AND MOCTA2.TA027=TD002 AND MOCTA2.TA028=TD003");
            SB.AppendFormat(@"  LEFT JOIN [TK].dbo.MOCTA MOCTA3 ON MOCTA3.TA021='02' AND MOCTA3.TA026=TD001 AND MOCTA3.TA027=TD002 AND MOCTA3.TA028=TD003");
            SB.AppendFormat(@"  LEFT JOIN [TK].dbo.MOCTA MOCTA4 ON MOCTA4.TA021='03' AND MOCTA4.TA026=TD001 AND MOCTA4.TA027=TD002 AND MOCTA4.TA028=TD003");
            SB.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(@"  AND COPTD.UDF01='Y' AND TD016='N' AND TD021='Y'");
            SB.AppendFormat(@"  AND (TD004 LIKE '4%' OR TD004 LIKE '5%')");
            SB.AppendFormat(@"  AND TD013>='{0}' AND  TD013<='{1}'", dateTimePicker15.Value.ToString("yyyyMMdd"), dateTimePicker16.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(@"  ORDER BY TC053,TD013,TD001,TD002,TD003");
            SB.AppendFormat(@"   ");
            SB.AppendFormat(@"   ");
            SB.AppendFormat(@"  ");


            return SB;

        }

        public void SETFASTREPORT4(string MANU,string SDAY,string EDAY)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL4(MANU, SDAY, EDAY);
            Report report1 = new Report();

            if(comboBox8.Text.Trim().Equals("1.總桶數"))
            {
                report1.Load(@"REPORT\預排製令矩陣-1.總桶數.frx");
            }
            else if (comboBox8.Text.Trim().Equals("2.包裝量"))
            {
                report1.Load(@"REPORT\預排製令矩陣-2.包裝量.frx");
            }
            else if (comboBox8.Text.Trim().Equals("3.數量+入庫量"))
            {
                report1.Load(@"REPORT\預排製令矩陣-3.數量+入庫量.frx");
            }
            else if (comboBox8.Text.Trim().Equals("4.包裝量+入庫量"))
            {
                report1.Load(@"REPORT\預排製令矩陣-4.包裝量+入庫量.frx");
            }
            else if (comboBox8.Text.Trim().Equals("5.桶數+數量"))
            {
                report1.Load(@"REPORT\預排製令矩陣-5.桶數+數量.frx");
            }
            else if (comboBox8.Text.Trim().Equals("6.桶數"))
            {
                report1.Load(@"REPORT\預排製令矩陣-6.桶數.frx");
            }
            else if (comboBox8.Text.Trim().Equals("A.商務組"))
            {
                report1.Load(@"REPORT\預排製令矩陣-A.商務組.frx");
            }
            else if (comboBox8.Text.Trim().Equals("B.商務組-包裝量+入庫量"))
            {
                report1.Load(@"REPORT\預排製令矩陣-B.商務組-包裝量+入庫量.frx");
            }
            else
            {
                report1.Load(@"REPORT\預排製令矩陣.frx");
            }


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

            report1.Preview = previewControl4;
            report1.Show();
        }

        public StringBuilder SETSQL4(string MANU, string SDAY, string EDAY)
        {
            StringBuilder Query = new StringBuilder();

            if(MANU.Equals("全部"))
            {
                Query.AppendFormat(@" ");
            }
            else
            {
                Query.AppendFormat(@" WHERE  TEMP.MANU='{0}'", MANU);
            }

            StringBuilder SB = new StringBuilder();

           
            SB.AppendFormat(@"    
                            SELECT MANU,MANUDATE,[MB002],BAR,NUM,PACKAGE,TD00123,TC053,MV002,MOCTA001002,入庫量,[NO],TA033,MOCTA001A,MOCTA002B,MOCTA001B,MOCTA002B
                            FROM (
                            SELECT  [MOCMANULINE].[MANU] ,CONVERT(nvarchar,[MOCMANULINE].[MANUDATE],112) MANUDATE,[MOCMANULINE].[MB002]
                            ,ISNULL([MOCMANULINE].[BAR],0) BAR,ISNULL([MOCMANULINE].[NUM],0) NUM,ISNULL([MOCMANULINE].[PACKAGE],0) PACKAGE
                            ,[MOCMANULINE].[COPTD001]+' '+[MOCMANULINE].[COPTD002]+' '+[MOCMANULINE].[COPTD003] AS TD00123
                            ,[COPTC].TC053,[CMSMV].MV002
                            ,ISNULL([MOCMANULINERESULT].[MOCTA001],'')+ISNULL([MOCMANULINERESULT].[MOCTA002],'')+ISNULL([MOCTA].TA001,'')+ISNULL([MOCTA].TA002,'') AS 'MOCTA001002' 
                            ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCMANULINERESULT].[MOCTA001] AND TG015=[MOCMANULINERESULT].[MOCTA002])+(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCTA].TA001 AND TG015=[MOCTA].TA002)  AS '入庫量'  
                            ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCMANULINERESULT].[MOCTA001] AND TG015=[MOCMANULINERESULT].[MOCTA002]) AS '入庫量A'  
                            ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCTA].TA001 AND TG015=[MOCTA].TA002)  AS '入庫量B'              
                            ,[MOCMANULINEMERGE].[NO],[MOCTA].TA033,ISNULL([MOCMANULINERESULT].[MOCTA001],'') AS MOCTA001A,ISNULL([MOCMANULINERESULT].[MOCTA002],'')  AS MOCTA002A,ISNULL([MOCTA].TA001,'')  AS MOCTA001B,ISNULL([MOCTA].TA002,'')  AS MOCTA002B  
                            FROM [TKMOC].[dbo].[MOCMANULINE]
                            LEFT JOIN [TK].dbo.[COPTD] ON [MOCMANULINE].[COPTD001]=[COPTD].TD001 AND [MOCMANULINE].[COPTD002]=[COPTD].TD002 AND[MOCMANULINE].[COPTD003]=[COPTD].TD003 
                            LEFT JOIN [TK].dbo.[COPTC] ON [COPTD].TD001=[COPTC].TC001 AND [COPTD].TD002=[COPTC].TC002
                            LEFT JOIN [TK].dbo.[CMSMV] ON [CMSMV].MV001=[COPTC].TC006
                            LEFT JOIN [TKMOC].[dbo].[MOCMANULINERESULT] ON [MOCMANULINERESULT].[SID]=[MOCMANULINE].[ID]
                            LEFT JOIN [TKMOC].[dbo].[MOCMANULINEMERGE] ON [MOCMANULINEMERGE].[SID]=[MOCMANULINE].[ID]  
                            LEFT JOIN [TK].dbo.[MOCTA] ON [MOCTA].TA033=[MOCMANULINEMERGE].[NO]  
                            WHERE CONVERT(nvarchar,[MOCMANULINE].[MANUDATE],112)>='{0}' AND CONVERT(nvarchar,[MOCMANULINE].[MANUDATE],112)<='{1}'
                            AND [MOCMANULINE].[MB001] NOT IN (SELECT MB001 FROM  [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                            UNION ALL  
                            SELECT  [MOCMANULINETEMP].[MANU] ,CONVERT(nvarchar,dateadd(ms,-3,dateadd(yy, datediff(yy,0,getdate())+2, 0)) ,112) MANUDATE,[MOCMANULINETEMP].[MB002]
                            ,ISNULL([MOCMANULINETEMP].[BAR],0) BAR,ISNULL([MOCMANULINETEMP].[NUM],0) NUM,ISNULL([MOCMANULINETEMP].[PACKAGE],0) PACKAGE
                            ,[MOCMANULINETEMP].[COPTD001]+' '+[MOCMANULINETEMP].[COPTD002]+' '+[MOCMANULINETEMP].[COPTD003] AS TD00123
                            ,[COPTC].TC053,[CMSMV].MV002
                            ,ISNULL([MOCMANULINERESULT].[MOCTA001],'')+ISNULL([MOCMANULINERESULT].[MOCTA002],'')+ISNULL([MOCTA].TA001,'')+ISNULL([MOCTA].TA002,'') AS 'MOCTA001002' 
                            ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCMANULINERESULT].[MOCTA001] AND TG015=[MOCMANULINERESULT].[MOCTA002])+(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCTA].TA001 AND TG015=[MOCTA].TA002)  AS '入庫量'  
                            ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCMANULINERESULT].[MOCTA001] AND TG015=[MOCMANULINERESULT].[MOCTA002]) AS '入庫量A'  
                            ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCTA].TA001 AND TG015=[MOCTA].TA002)  AS '入庫量B'  
                            ,[MOCMANULINEMERGE].[NO],[MOCTA].TA033,ISNULL([MOCMANULINERESULT].[MOCTA001],'') AS MOCTA001A,ISNULL([MOCMANULINERESULT].[MOCTA002],'')  AS MOCTA002A,ISNULL([MOCTA].TA001,'')  AS MOCTA001B,ISNULL([MOCTA].TA002,'')  AS MOCTA002B  
                            FROM [TKMOC].[dbo].[MOCMANULINETEMP]  
                            LEFT JOIN [TK].dbo.[COPTD] ON [MOCMANULINETEMP].[COPTD001]=[COPTD].TD001 AND [MOCMANULINETEMP].[COPTD002]=[COPTD].TD002 AND[MOCMANULINETEMP].[COPTD003]=[COPTD].TD003   
                            LEFT JOIN [TK].dbo.[COPTC] ON [COPTD].TD001=[COPTC].TC001 AND [COPTD].TD002=[COPTC].TC002  
                            LEFT JOIN [TK].dbo.[CMSMV] ON [CMSMV].MV001=[COPTC].TC006  
                            LEFT JOIN [TKMOC].[dbo].[MOCMANULINE] ON [MOCMANULINE].ID=[MOCMANULINETEMP].TID  
                            LEFT JOIN [TKMOC].[dbo].[MOCMANULINERESULT] ON [MOCMANULINERESULT].[SID]=[MOCMANULINE].[ID]  
                            LEFT JOIN [TKMOC].[dbo].[MOCMANULINEMERGE] ON [MOCMANULINEMERGE].[SID]=[MOCMANULINE].[ID]  
                            LEFT JOIN [TK].dbo.[MOCTA] ON [MOCTA].TA033=[MOCMANULINEMERGE].[NO]  
                            WHERE CONVERT(nvarchar,[MOCMANULINETEMP].[MANUDATE],112)>='{2}' AND CONVERT(nvarchar,[MOCMANULINETEMP].[MANUDATE],112)<='{3}' 
                            AND [MOCMANULINETEMP].TID IS NULL  
                            AND [MOCMANULINE].[MB001] NOT IN (SELECT MB001 FROM  [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                            ) AS TEMP
                            {4}
                            ORDER BY  TEMP.[MANU],CONVERT(nvarchar, TEMP.[MANUDATE],112)
                            ", SDAY, EDAY, SDAY, EDAY, Query.ToString());


            return SB;

        }
        public void ADDTOUOFTB_EIP_SCH_MEMO(string SDay,string EDay)
        {
            DataSet ds = new DataSet();
            ds=SEARCHMANULINE(SDay, EDay);

            try
            {
                
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

               

                //[CREATE_USER]='3edfae1f-f607-4651-b351-6df86725d5a8'，包裝線 MANU90
                //[CREATE_USER]='6c4bcbe1-52ed-4277-8259-4d765ca529b7'，製一線 MANU10
                //[CREATE_USER]='cfd7fa21-c3b4-4174-81ff-b0b92aa9ab9c'，製二線 MANU20
                //[CREATE_USER]='7efba0e4-220f-4843-93de-37f05e370fcf'，手工線 MANU30
                //將資料從TKMOC的MOCMANULINE計算出工時，再COPY到UOF的TB_EIP_SCH_MEMO
                //先刪除再新增

                sbSql.AppendFormat(" DELETE [UOFTEST].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND  CONVERT(NVARCHAR,[START_TIME],112) <='{1}' AND [CREATE_USER]='3edfae1f-f607-4651-b351-6df86725d5a8'", SDay, EDay);
                sbSql.AppendFormat(" DELETE [UOFTEST].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND  CONVERT(NVARCHAR,[START_TIME],112) <='{1}' AND [CREATE_USER]='6c4bcbe1-52ed-4277-8259-4d765ca529b7'", SDay, EDay);
                sbSql.AppendFormat(" DELETE [UOFTEST].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND  CONVERT(NVARCHAR,[START_TIME],112) <='{1}' AND [CREATE_USER]='cfd7fa21-c3b4-4174-81ff-b0b92aa9ab9c'", SDay, EDay);
                sbSql.AppendFormat(" DELETE [UOFTEST].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND  CONVERT(NVARCHAR,[START_TIME],112) <='{1}' AND [CREATE_USER]='7efba0e4-220f-4843-93de-37f05e370fcf'", SDay, EDay);
                sbSql.AppendFormat(" ");

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    sbSql.AppendFormat(" INSERT INTO [UOFTEST].[dbo].[TB_EIP_SCH_MEMO]");
                    sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                    sbSql.AppendFormat(" VALUES");
                    sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                    sbSql.AppendFormat(" ");
                }


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

        public DataSet SEARCHMANULINE(string Sday,string EDay)
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

                sbSql.AppendFormat(" SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'3edfae1f-f607-4651-b351-6df86725d5a8' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '+'架動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/16*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112) AS [START_TIME],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '+'架動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/16*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'3edfae1f-f607-4651-b351-6df86725d5a8' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='包裝線'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'3edfae1f-f607-4651-b351-6df86725d5a8' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'3edfae1f-f607-4651-b351-6df86725d5a8' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='包裝線'");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'6c4bcbe1-52ed-4277-8259-4d765ca529b7' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做24桶 '+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' +'架動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/14*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112) AS [START_TIME],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做24桶 '+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '+'架動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/14*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'6c4bcbe1-52ed-4277-8259-4d765ca529b7' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='製一線'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) ");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'6c4bcbe1-52ed-4277-8259-4d765ca529b7' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'6c4bcbe1-52ed-4277-8259-4d765ca529b7' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='製一線'");
                 sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'cfd7fa21-c3b4-4174-81ff-b0b92aa9ab9c' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做38桶 '+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '+'架動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/14*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112) AS [START_TIME],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做38桶 '+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '+'架動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/14*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'cfd7fa21-c3b4-4174-81ff-b0b92aa9ab9c' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='製二線'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'cfd7fa21-c3b4-4174-81ff-b0b92aa9ab9c' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'cfd7fa21-c3b4-4174-81ff-b0b92aa9ab9c' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='製二線'");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7efba0e4-220f-4843-93de-37f05e370fcf' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '+'架動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/6.5*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112) AS [START_TIME],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '+'架動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/6.5*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7efba0e4-220f-4843-93de-37f05e370fcf' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='手工線'");
                sbSql.AppendFormat(" GROUP BY [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) ");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7efba0e4-220f-4843-93de-37f05e370fcf' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7efba0e4-220f-4843-93de-37f05e370fcf' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='手工線'");     
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat(" ORDER BY [START_TIME]");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter8 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder8 = new SqlCommandBuilder(adapter8);
                sqlConn.Open();
                ds8.Clear();
                adapter8.Fill(ds8, "ds8");
              


                if (ds8.Tables["ds8"].Rows.Count == 0)
                {
                    return ds8;
                }
                else
                {
                    if (ds8.Tables["ds8"].Rows.Count >= 1)
                    {
                        return ds8;
                    }

                    return ds8;
                }

            }
            catch
            {
                return ds8;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDTOUOFTB_EIP_SCH_MEMO_MOC(string Sday)
        {

            DataSet ds = new DataSet();
            ds = SEARCHMANULINE(Sday);
            Thread.Sleep(1000);
            ds2 = SEARCHMANULINE2(Sday);
            Thread.Sleep(1000);
            ds3 = SEARCHMANULINE3(Sday);
            Thread.Sleep(1000);
            ds4 = SEARCHMANULINE4(Sday);

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);


                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                //[CREATE_USER]='7774b96c-6762-45ef-b9d1-fcd718854e9f'，包裝線 MANU90
                //[CREATE_USER]='5ce0f554-8b80-4aed-afea-fcd224cecb81'，製一線 MANU10
                //[CREATE_USER]='0c98530a-b467-4cd4-a411-7279f1e04d0d'，製二線 MANU20
                //[CREATE_USER]='88789ece-41d1-4b48-94f1-6ffab05b05f4'，手工線 MANU30
                //將資料從TKMOC的MOCMANULINE計算出工時，再COPY到UOF的TB_EIP_SCH_MEMO
                //先刪除再新增

                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='7774b96c-6762-45ef-b9d1-fcd718854e9f'", Sday);
                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='5ce0f554-8b80-4aed-afea-fcd224cecb81'", Sday);
                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='0c98530a-b467-4cd4-a411-7279f1e04d0d'", Sday);
                sbSql.AppendFormat(" DELETE [UOF].[dbo].[TB_EIP_SCH_MEMO] WHERE CONVERT(NVARCHAR,[START_TIME],112)>='{0}' AND [CREATE_USER]='88789ece-41d1-4b48-94f1-6ffab05b05f4'", Sday);
                sbSql.AppendFormat(" ");

                if (ds11.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }

                if (ds12.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds2.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }

                if (ds13.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds3.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }

                if (ds14.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds4.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [UOF].[dbo].[TB_EIP_SCH_MEMO]");
                        sbSql.AppendFormat(" ([CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID])");
                        sbSql.AppendFormat(" VALUES");
                        sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", dr["CREATE_TIME"].ToString(), dr["CREATE_USER"].ToString(), dr["DESCRIPTION"].ToString(), dr["END_TIME"].ToString(), dr["MEMO_GUID"].ToString(), dr["PERSONAL_TYPE"].ToString(), dr["REPEAT_GUID"].ToString(), dr["START_TIME"].ToString(), dr["SUBJECT"].ToString(), dr["REMINDER_GUID"].ToString(), dr["ALL_DAY"].ToString(), dr["OWNER"].ToString(), dr["UID"].ToString(), dr["ICS_GUID"].ToString());
                        sbSql.AppendFormat(" ");
                    }
                }




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

        //製一線、製二線的桶數
        public DataSet SEARCHMANULINE(string Sday)
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



                sbSql.AppendFormat(@" 
                                    SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]
                                    FROM (
                                    SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{1}桶 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,1,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'---'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{1}桶 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                    FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                    LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'
                                    WHERE INVMB.MB001=MOCMANULINE.MB001 
                                    AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}'
                                    AND [MOCMANULINE]. [MANU]='製一線'
                                    AND [MOCMANULINE].[MB001] NOT IN (SELECT MB001 FROM  [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])      
                                    GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                    UNION
                                    SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{2}桶 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,1,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'---'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),SUM([BAR])))+'桶數/每日可做{2}桶 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                    FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                    LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'
                                    WHERE INVMB.MB001=MOCMANULINE.MB001   
                                    AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}'
                                    AND [MOCMANULINE]. [MANU]='製二線'
                                    AND [MOCMANULINE].[MB001] NOT IN (SELECT MB001 FROM  [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                                    GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                    ) AS TEMP
                                    ORDER BY [START_TIME],[SUBJECT]
                                    ", DateTime.Now.ToString("yyyyMMdd"), BASELIMITHRSBAR1, BASELIMITHRSBAR2);

                adapter11 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder11 = new SqlCommandBuilder(adapter11);
                sqlConn.Open();
                ds11.Clear();
                adapter11.Fill(ds11, "ds11");



                if (ds11.Tables["ds11"].Rows.Count == 0)
                {
                    return ds11;
                }
                else
                {
                    if (ds11.Tables["ds11"].Rows.Count >= 1)
                    {
                        return ds11;
                    }

                    return ds11;
                }

            }
            catch
            {
                return ds11;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        //包裝線、製一線、製二線、手工線的總工時
        public DataSet SEARCHMANULINE2(string Sday)
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
                sbSql.AppendFormat(@" 
                                     SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]
                                     FROM (
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001   
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='包裝線'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]                
                                     UNION
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 '  AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)  AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001 
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='製一線'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                     UNION
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)  AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001   
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='製二線'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                     UNION               
                                     SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,2,[MANUDATE]),21)   AS [START_TIME],[MOCMANULINE].[MANU]+'--總工時-'+CONVERT(nvarchar,SUM(CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))))+'小時 ' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                     FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                     LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'
                                     WHERE INVMB.MB001=MOCMANULINE.MB001 
                                     AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                     AND [MOCMANULINE]. [MANU]='手工線'
                                     GROUP BY [MOCMANULINE].[MANU],[MANUDATE]              
                                     ) AS TEMP
                                     ORDER BY [START_TIME],[SUBJECT]
                                    ", DateTime.Now.ToString("yyyyMMdd"));

                adapter12 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder12 = new SqlCommandBuilder(adapter12);
                sqlConn.Open();
                ds12.Clear();
                adapter12.Fill(ds12, "ds12");



                if (ds12.Tables["ds12"].Rows.Count == 0)
                {
                    return ds12;
                }
                else
                {
                    if (ds12.Tables["ds12"].Rows.Count >= 1)
                    {
                        return ds12;
                    }

                    return ds12;
                }

            }
            catch
            {
                return ds12;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        //包裝線、製一線、製二線、手工線的稼動率
        public DataSet SEARCHMANULINE3(string Sday)
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

                sbSql.AppendFormat(@" 
                                 SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]
                                 FROM (               
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{4}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{4}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001   
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='包裝線'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                 UNION
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{1}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{1}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001 
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='製一線'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE] 
                                 UNION
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{2}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{2}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001   
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='製二線'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE]
                                 UNION                
                                 SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{3}*100))+'%' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(varchar(100),DATEADD(second,3,[MANUDATE]),21) AS [START_TIME],[MOCMANULINE].[MANU]+'-稼動率-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),SUM(ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0))/{3}*100))+'%' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]
                                 FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB
                                 LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'
                                 WHERE INVMB.MB001=MOCMANULINE.MB001 
                                 AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' 
                                 AND [MOCMANULINE]. [MANU]='手工線'
                                 GROUP BY [MOCMANULINE].[MANU],[MANUDATE]                
                                 ) AS TEMP
                                 ORDER BY [START_TIME],[SUBJECT]
                                 ", DateTime.Now.ToString("yyyyMMdd"), BASELIMITHRS1, BASELIMITHRS2, BASELIMITHRS3, BASELIMITHRS9);


                adapter13 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder13 = new SqlCommandBuilder(adapter13);
                sqlConn.Open();
                ds13.Clear();
                adapter13.Fill(ds13, "ds13");



                if (ds13.Tables["ds13"].Rows.Count == 0)
                {
                    return ds13;
                }
                else
                {
                    if (ds13.Tables["ds13"].Rows.Count >= 1)
                    {
                        return ds13;
                    }

                    return ds13;
                }

            }
            catch
            {
                return ds13;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        //包裝線、製一線、製二線、手工線的明細
        public DataSet SEARCHMANULINE4(string Sday)
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

                sbSql.AppendFormat(" SELECT [CREATE_TIME],[CREATE_USER],[DESCRIPTION],[END_TIME],[MEMO_GUID],[PERSONAL_TYPE],[REPEAT_GUID],[START_TIME],[SUBJECT],[REMINDER_GUID],[ALL_DAY],[OWNER],[UID],[ICS_GUID]");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'7774b96c-6762-45ef-b9d1-fcd718854e9f' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='包裝線'");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'0c98530a-b467-4cd4-a411-7279f1e04d0d' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製一線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='製一線'");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(NVARCHAR,CONVERT(DECIMAL(14,2),[BAR]))+'桶數-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'5ce0f554-8b80-4aed-afea-fcd224cecb81' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%製二線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001 ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='製二線'");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [DESCRIPTION],CONVERT(NVARCHAR,[MANUDATE],112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,[MANUDATE],112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],[MOCMANULINE].[MANU]+[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[NUM]))+INVMB.MB004+'-'+CONVERT(nvarchar,CONVERT(DECIMAL(12,2),ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0)))+'小時' AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                sbSql.AppendFormat(" LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%手工線%'");
                sbSql.AppendFormat(" WHERE INVMB.MB001=MOCMANULINE.MB001   ");
                sbSql.AppendFormat(" AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' ", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" AND [MOCMANULINE]. [MANU]='手工線'");
                sbSql.AppendFormat(" UNION");
                sbSql.AppendFormat(" SELECT CONVERT(varchar(100),GETDATE(),21) AS [CREATE_TIME],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [CREATE_USER],'手工線'+TA001+'-'+TA002+TA034+CONVERT(NVARCHAR,CONVERT(INT,TA015))+TA007 AS [DESCRIPTION],CONVERT(NVARCHAR,TA003,112) AS [END_TIME],NEWID() AS [MEMO_GUID],'Display' AS [PERSONAL_TYPE],NULL AS [REPEAT_GUID],CONVERT(NVARCHAR,TA003,112)+' '+CONVERT(varchar(100),GETDATE(),14) AS [START_TIME],'手工線'+TA001+'-'+TA002+TA034+CONVERT(NVARCHAR,CONVERT(INT,TA015))+TA007 AS [SUBJECT],NULL AS [REMINDER_GUID],'1' AS [ALL_DAY],'88789ece-41d1-4b48-94f1-6ffab05b05f4' AS [OWNER],NULL AS [UID],NULL AS [ICS_GUID]");
                sbSql.AppendFormat(" FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(" WHERE  TA021='04'");
                sbSql.AppendFormat(" AND TA003>='{0}'", DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat(" ORDER BY [START_TIME],[SUBJECT]");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                adapter14 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder14 = new SqlCommandBuilder(adapter14);
                sqlConn.Open();
                ds14.Clear();
                adapter14.Fill(ds14, "ds14");



                if (ds14.Tables["ds14"].Rows.Count == 0)
                {
                    return ds14;
                }
                else
                {
                    if (ds14.Tables["ds14"].Rows.Count >= 1)
                    {
                        return ds14;
                    }

                    return ds14;
                }

            }
            catch
            {
                return ds14;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public decimal SEARCHBASELIMITHRS(string ID)
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


                sbSql.AppendFormat(@" 
                                    SELECT  [ID],[LIMITHRS] FROM [TKMOC].[dbo].[BASELIMITHRS] WHERE [ID]='{0}'
                                    ", ID);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToDecimal(ds1.Tables["ds1"].Rows[0]["LIMITHRS"].ToString());
                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
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
            Search();
            //SEARCHMOCMANULINECOP();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SearchMATRIAL();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ExcelExportMATERIAL();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SearchV2();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SEARCHMOCTG();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        private void button42_Click(object sender, EventArgs e)
        {
            SEARCHCOPTD();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            RESET();
            SETPATH();
            SETFILE();

            MessageBox.Show("OK");
        }
        private void button9_Click(object sender, EventArgs e)
        {
            CLEAREXCEL();
        }



        private void button11_Click(object sender, EventArgs e)
        {
            RESET2();
            SETPATH2();
            SETFILE2();

            MessageBox.Show("OK");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            CLEAREXCEL();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        private void button13_Click(object sender, EventArgs e)
        {
            RESET3();
            SETPATH3();
            SETFILE3();

            MessageBox.Show("OK");
        }
        private void button14_Click(object sender, EventArgs e)
        {
            CLEAREXCEL();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }
        private void button16_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3();
        }
        private void button17_Click(object sender, EventArgs e)
        {
            RESET4();
            SETPATH4();
            SETFILE4();

            MessageBox.Show("OK");
        }
        private void button19_Click(object sender, EventArgs e)
        {
            //ADDTOUOFTB_EIP_SCH_MEMO(dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));

            BASELIMITHRSBAR1 = SEARCHBASELIMITHRS("製一線桶數");
            BASELIMITHRSBAR1 = Math.Round(BASELIMITHRSBAR1, 0);
            BASELIMITHRSBAR2 = SEARCHBASELIMITHRS("製二線桶數");
            BASELIMITHRSBAR2 = Math.Round(BASELIMITHRSBAR2, 0);

            BASELIMITHRS1 = SEARCHBASELIMITHRS("製一線稼動率時數");
            BASELIMITHRS2 = SEARCHBASELIMITHRS("製二線稼動率時數");
            BASELIMITHRS3 = SEARCHBASELIMITHRS("手工線稼動率時數");
            BASELIMITHRS9 = SEARCHBASELIMITHRS("包裝線稼動率時數");

            ADDTOUOFTB_EIP_SCH_MEMO_MOC(dateTimePicker11.Value.ToString("yyyyMMdd"));
        }

        private void button20_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4(comboBox7.Text,dateTimePicker19.Value.ToString("yyyyMMdd"), dateTimePicker20.Value.ToString("yyyyMMdd"));
        }

        #endregion


    }
}
