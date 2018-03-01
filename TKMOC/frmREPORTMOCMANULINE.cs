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
        DataSet dsCALENDAR = new DataSet();

        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds22 = new DataSet();
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        string tablename = null;
        int rownum = 0;

        string SOURCEID;

        public frmREPORTMOCMANULINE()
        {
            InitializeComponent();

            SETCALENDAR();

            //comboBox1load();
        }

        #region FUNCTION
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
                STR.AppendFormat(@"  ,[MOCMANULINE].CLINET AS '客戶',ISNULL([MOCMANULINE].MANUHOUR,0) AS '生產時數'");
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
                STR.AppendFormat(@"  ,[MOCMANULINE].CLINET AS '客戶',ISNULL([MOCMANULINE].MANUHOUR,0) AS '生產時數'");
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

        #endregion

       
    }
}
