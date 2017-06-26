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
    public partial class frmREPORTEngineering : Form
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


        public frmREPORTEngineering()
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
           
            string ThisYear = null;
            string ThisMonth = null;
            string LastMonth = null;
            string LastYear = null;
            string LastYearMonth = null;

          

            if (comboBox1.Text.ToString().Equals("共用性備品"))
            {

                
                STR.AppendFormat(@"  SELECT [EQUIPMENTID] AS '備品編號',[EQUIPMENTNAME] AS '品名',[SPEC] AS '規格'");
                STR.AppendFormat(@"  ,[PRICES] AS '單價',[SUPPLY] AS '供應商',[TEL] AS '電話'");
                STR.AppendFormat(@"  ,[USEDTIME] AS '使用壽命',[SAFENUM] AS '安全庫存'");
                STR.AppendFormat(@"  ,[BUYIN] AS '已進數量',[USED] AS '已用數量',[NOWS] AS '可用數量' ,[PRETIME] AS '前置時間'");
                STR.AppendFormat(@"  ,[EQUIPMENT] AS '適用設備'");
                STR.AppendFormat(@"  ,[COMMENT] AS '備註'");
                STR.AppendFormat(@"  ,[ID]");
                STR.AppendFormat(@"  FROM [TKMOC].[dbo].[MACHINECOMMON]");
                STR.AppendFormat(@"  ORDER BY [EQUIPMENTID]");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds1";
            }
            
            else if (comboBox1.Text.ToString().Equals(""))
            {

            }


            return STR;
        }

        public void ExcelExport()
        {
            Search();
            string TABLENAME = "報表";

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
                TABLENAME = "共用性備品";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds1"].Rows[i][rows].ToString());
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

        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
    }
}
