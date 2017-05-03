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
    public partial class frmREPORTMOCMANULINE : Form
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


        public frmREPORTMOCMANULINE()
        {
            InitializeComponent();

            comboBox1load();
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
       

            if (!string.IsNullOrEmpty(comboBox1.Text.ToString()))
            {                
                STR.AppendFormat(@"  SELECT [MOCMANULINE].MANU AS '線別',CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) AS '生產日'");
                STR.AppendFormat(@"  ,[MOCMANULINE].MB001 AS '品號',[MOCMANULINE].MB002 AS '品名',[MOCMANULINE].MB003 AS '規格'");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINE].BAR,0) AS '桶數',ISNULL([MOCMANULINE].NUM,0) AS '片數',ISNULL([MOCMANULINE].BOX,0) AS '箱數',ISNULL([MOCMANULINE].PACKAGE,0) AS '包裝數'");
                STR.AppendFormat(@"  ,ISNULL(CONVERT(NVARCHAR(10),[MOCMANULINE].OUTDATE,112),'') AS '預交日'");
                STR.AppendFormat(@"  ,[MOCMANULINE].CLINET AS '客戶',ISNULL([MOCMANULINE].MANUHOUR,0) AS '生產時數'");
                STR.AppendFormat(@"  ,ISNULL([MOCMANULINERESULT].MOCTA001,'') AS '製令單別',ISNULL([MOCMANULINERESULT].MOCTA002,'') AS '製令單號'");
                STR.AppendFormat(@"  FROM [TKMOC].dbo.[MOCMANULINE]");
                STR.AppendFormat(@"  LEFT JOIN  [TKMOC].dbo.[MOCMANULINERESULT] ON [MOCMANULINE].ID=[MOCMANULINERESULT].SID");
                STR.AppendFormat(@"  WHERE [MOCMANULINE].MANU='{0}'", comboBox1.Text.ToString());
                STR.AppendFormat(@"  AND CONVERT(NVARCHAR(10),[MOCMANULINE].MANUDATE,112) LIKE '{0}%'",dateTimePicker1.Value.ToString("yyyyMM"));
                STR.AppendFormat(@"  ");


                tablename = "TEMPds1";
            }
           
            else if (comboBox1.Text.ToString().Equals(""))
            {
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds2";
            }
            



            return STR;
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        #endregion


    }
}
