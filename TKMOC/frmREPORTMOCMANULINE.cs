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
        DataSet dsCALENDAR = new DataSet();

        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();

        DataSet ds2 = new DataSet();
        DataSet ds22 = new DataSet();
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();

        string tablename = null;
        int rownum = 0;

        string SOURCEID;
        string DATES = null;
        string strDesktopPath;
        string pathFile;
        string pathFile2;

        string[] message = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] message2 = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        DateTime sdt;
        DateTime edt;
        DateTime sdt2;
        DateTime edt2;

        public frmREPORTMOCMANULINE()
        {
            InitializeComponent();

            SETCALENDAR();

            //comboBox1load();

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
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'   ");
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
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if(comboBox4.Text.Equals("新廠包裝線"))
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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder SB = new StringBuilder();

                if (comboBox6.Text.Equals("新廠包裝線"))
                {
                    sbSql.AppendFormat(@"  SELECT  CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)+' '+[MOCMANULINE].[MANU] AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004 AS ' PACKAGE'");
                    sbSql.AppendFormat(@"  ,'---' AS '-'");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                    sbSql.AppendFormat(@"  WHERE INVMB.MB001=MOCMANULINE.MB001    ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='20200201' AND CONVERT(NVARCHAR,[MANUDATE],112) <='20200228' ");
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='新廠包裝線'");
                    sbSql.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU],[MANUDATE]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox6.Text.Equals("新廠製二組"))
                {
                    sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)+' '+[MOCMANULINE].[MANU] AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS ' MB002', CONVERT(NVARCHAR,CONVERT(INT,ROUND([NUM],0)))++' KG' AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,'---' AS '-' ");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB ");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製二組%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004            ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='新廠製二組'");
                    sbSql.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox6.Text.Equals("新廠製一組"))
                {
                    sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)+' '+[MOCMANULINE].[MANU] AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS ' MB002', CONVERT(NVARCHAR,CONVERT(INT,ROUND([NUM],0)))++' KG' AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,'---' AS '-' ");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB ");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製一組%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004            ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='新廠製一組'");
                    sbSql.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                }
                else if (comboBox6.Text.Equals("新廠製三組(手工)"))
                {
                    sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112)+' '+[MOCMANULINE].[MANU] AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS ' MB002', CONVERT(NVARCHAR,CONVERT(INT,ROUND([NUM],0)))++' KG' AS 'PACKAGE'");
                    sbSql.AppendFormat(@"  ,'---' AS '-' ");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB ");
                    sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製三組(手工)%'");
                    sbSql.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                    sbSql.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004            ");
                    sbSql.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='新廠製三組(手工)'");
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

        public void ExportDataSetToExcel2(DataSet ds, string TopathFile)
        {
            int days = Convert.ToInt32(sdt.AddDays(-sdt.Day + 1).DayOfWeek.ToString("d"));
            //MessageBox.Show(days.ToString());
            int MONTHDAYS = DateTime.DaysInMonth(sdt.Year, sdt.Month);

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
                            message2[0] = message2[0] + table.Rows[j].ItemArray[k].ToString();
                            message2[0] = message2[0] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 2)
                        {
                            message2[1] = message2[1] + table.Rows[j].ItemArray[k].ToString();
                            message2[1] = message2[1] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 3)
                        {
                            message2[2] = message2[2] + table.Rows[j].ItemArray[k].ToString();
                            message2[2] = message2[2] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 4)
                        {
                            message2[3] = message2[3] + table.Rows[j].ItemArray[k].ToString();
                            message2[3] = message2[3] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 5)
                        {
                            message2[4] = message2[4] + table.Rows[j].ItemArray[k].ToString();
                            message2[4] = message2[4] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 6)
                        {
                            message2[5] = message2[5] + table.Rows[j].ItemArray[k].ToString();
                            message2[5] = message2[5] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 7)
                        {
                            message2[6] = message2[6] + table.Rows[j].ItemArray[k].ToString();
                            message2[6] = message2[6] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 8)
                        {
                            message2[7] = message2[7] + table.Rows[j].ItemArray[k].ToString();
                            message2[7] = message2[7] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 9)
                        {
                            message2[8] = message2[8] + table.Rows[j].ItemArray[k].ToString();
                            message2[8] = message2[8] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 10)
                        {
                            message2[9] = message2[9] + table.Rows[j].ItemArray[k].ToString();
                            message2[9] = message2[9] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 11)
                        {
                            message2[10] = message2[10] + table.Rows[j].ItemArray[k].ToString();
                            message2[10] = message2[10] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 12)
                        {
                            message2[11] = message2[11] + table.Rows[j].ItemArray[k].ToString();
                            message2[11] = message2[11] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 13)
                        {
                            message2[12] = message2[12] + table.Rows[j].ItemArray[k].ToString();
                            message2[12] = message2[12] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 14)
                        {
                            message2[13] = message2[13] + table.Rows[j].ItemArray[k].ToString();
                            message2[13] = message2[13] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 15)
                        {
                            message2[14] = message2[14] + table.Rows[j].ItemArray[k].ToString();
                            message2[14] = message2[14] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 16)
                        {
                            message2[15] = message2[15] + table.Rows[j].ItemArray[k].ToString();
                            message2[15] = message2[15] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 17)
                        {
                            message2[16] = message2[16] + table.Rows[j].ItemArray[k].ToString();
                            message2[16] = message2[16] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 18)
                        {
                            message2[17] = message2[17] + table.Rows[j].ItemArray[k].ToString();
                            message2[17] = message2[17] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 19)
                        {
                            message2[18] = message2[18] + table.Rows[j].ItemArray[k].ToString();
                            message2[18] = message2[18] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 20)
                        {
                            message2[19] = message2[19] + table.Rows[j].ItemArray[k].ToString();
                            message2[19] = message2[19] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 21)
                        {
                            message2[20] = message2[20] + table.Rows[j].ItemArray[k].ToString();
                            message2[20] = message2[20] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 22)
                        {
                            message2[21] = message2[21] + table.Rows[j].ItemArray[k].ToString();
                            message2[21] = message2[21] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 23)
                        {
                            message2[22] = message2[22] + table.Rows[j].ItemArray[k].ToString();
                            message2[22] = message2[22] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 24)
                        {
                            message2[23] = message2[23] + table.Rows[j].ItemArray[k].ToString();
                            message2[23] = message2[23] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 25)
                        {
                            message2[24] = message2[24] + table.Rows[j].ItemArray[k].ToString();
                            message2[24] = message2[24] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 26)
                        {
                            message2[25] = message2[25] + table.Rows[j].ItemArray[k].ToString();
                            message2[25] = message2[25] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 27)
                        {
                            message2[26] = message2[26] + table.Rows[j].ItemArray[k].ToString();
                            message2[26] = message2[26] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 28)
                        {
                            message2[27] = message2[27] + table.Rows[j].ItemArray[k].ToString();
                            message2[27] = message2[27] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 29)
                        {
                            message2[28] = message2[28] + table.Rows[j].ItemArray[k].ToString();
                            message2[28] = message2[28] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 30)
                        {
                            message2[29] = message2[29] + table.Rows[j].ItemArray[k].ToString();
                            message2[29] = message2[29] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 31)
                        {
                            message2[30] = message2[30] + table.Rows[j].ItemArray[k].ToString();
                            message2[30] = message2[30] + '\n';
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

                    excelWorkSheet.Cells[EXCELX, EXCELY] = message2[i - 1].ToString();

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

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            if (comboBox6.Text.Equals("新廠包裝線"))
            {
                SB.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002]+' '+CONVERT(NVARCHAR,CONVERT(INT,ROUND([MOCMANULINE].[BOX],0)))+' 箱 '+CONVERT(NVARCHAR,CONVERT(INT,[MOCMANULINE].[PACKAGE]))+INVMB.MB004 AS 'PACKAGE'");
                SB.AppendFormat(@"  ,ISNULL(ROUND(([PACKAGE]/NULLIF([PREINVMBMANU].TIMES,1)),0),0) AS HRS");
                SB.AppendFormat(@"  ,INVMB.MB001");
                SB.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.INVMB");
                SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%包裝%'");
                SB.AppendFormat(@"  WHERE INVMB.MB001=MOCMANULINE.MB001      ");
                SB.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ",dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='新廠包裝線'");
                SB.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU],[MANUDATE]");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
            }
            else if (comboBox6.Text.Equals("新廠製二組"))
            {
                SB.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS 'PACKAGE'");
                SB.AppendFormat(@"  ,ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0) AS HRS");
                SB.AppendFormat(@"  ,INVMB.MB001");
                SB.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB");
                SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製二組%'");
                SB.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                SB.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                SB.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='新廠製二組'");
                SB.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
            }
            else if (comboBox6.Text.Equals("新廠製一組"))
            {
                SB.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS 'PACKAGE'");
                SB.AppendFormat(@"  ,ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0) AS HRS");
                SB.AppendFormat(@"  ,INVMB.MB001");
                SB.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB");
                SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製一組%'");
                SB.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                SB.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                SB.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='新廠製一組'");
                SB.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
            }
            else if (comboBox6.Text.Equals("新廠製三組(手工)"))
            {
                SB.AppendFormat(@"  SELECT  [MOCMANULINE].[MANU],CONVERT(NVARCHAR,[MOCMANULINE].[MANUDATE],112) AS MANUDATE ,[MOCMANULINE].[COPTD001]+'-'+[MOCMANULINE].[COPTD002]+'-'+[MOCMANULINE].[COPTD003]+'-'+INVMB.[MB002] AS 'PACKAGE'");
                SB.AppendFormat(@"  ,ISNULL(ROUND(([NUM]/NULLIF([PREINVMBMANU].TIMES,1)),0),0) AS HRS");
                SB.AppendFormat(@"  ,INVMB.MB001");
                SB.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].dbo.COPTD,[TK].dbo.INVMB");
                SB.AppendFormat(@"  LEFT JOIN [TKMOC].[dbo].[PREINVMBMANU] ON [PREINVMBMANU].MB001=INVMB.MB001 AND [PREINVMBMANU].MANU LIKE '%新廠製三組(手工)%'");
                SB.AppendFormat(@"  WHERE [MOCMANULINE].COPTD001=COPTD.TD001 AND [MOCMANULINE].COPTD002=COPTD.TD002 AND [MOCMANULINE].COPTD003=COPTD.TD003");
                SB.AppendFormat(@"  AND INVMB.MB001=COPTD.TD004");
                SB.AppendFormat(@"  AND CONVERT(NVARCHAR,[MANUDATE],112) >='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112) <='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND [MOCMANULINE]. [MANU]='新廠製三組(手工)'");
                SB.AppendFormat(@"  ORDER BY [MOCMANULINE].[MANU], [MOCMANULINE].[MANUDATE]");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
            }


            return SB;

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
            RESET();
            SETPATH3();
            SETFILE3();

            MessageBox.Show("OK");
        }
        private void button14_Click(object sender, EventArgs e)
        {
            CLEAREXCEL();
        }

        #endregion



    }
}
