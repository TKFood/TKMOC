using System;
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
using System.Text.RegularExpressions;

namespace TKMOC
{
    public partial class frmMOCPLANWEEK : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        DataTable dtTemp = new DataTable();
        DataTable dtTemp2 = new DataTable();

        string tablename = null;
        decimal COPNum = 0;
        decimal TOTALCOPNum = 0;
        double BOMNum = 0;
        double FinalNum = 0;
        decimal COOKIES = 1;
        decimal BATCH = 1;
        Thread TD;

        public frmMOCPLANWEEK()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            StringBuilder TD001 = new StringBuilder();
            StringBuilder TC027 = new StringBuilder();

            if (checkBox1.Checked == true)
            {
                TD001.Append("'A221',");
            }
            if (checkBox2.Checked == true)
            {
                TD001.Append("'A222',");
            }

            if (checkBox4.Checked == true)
            {
                TD001.Append("'A225',");
            }
            if (checkBox5.Checked == true)
            {
                TD001.Append("'A226',");
            }
            if (checkBox6.Checked == true)
            {
                TD001.Append("'A227',");
            }
            if (checkBox7.Checked == true)
            {
                TD001.Append("'A223',");
            }

            if (comboBox1.Text.ToString().Equals("已確認"))
            {
                TC027.Append(" AND TC027='Y' ");
            }
            else if (comboBox1.Text.ToString().Equals("未確認(扣已確認)"))
            {
                TC027.Append("AND TC027='N' ");
            }
            else if (comboBox1.Text.ToString().Equals("全部"))
            {
                TC027.Append("  ");
            }
            TD001.Append("''");

            dtTemp.Clear();
            dtTemp2.Clear();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@"  SELECT 客戶,日期,品號,品名,規格,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位 ,單別,單號,序號  ");
                sbSql.Append(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20001' AND LA001=品號) AS '成品倉庫存'");
                sbSql.Append(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20002' AND LA001=品號) AS '外銷倉庫存'");
                sbSql.Append(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(TA015-TA017-TA018),0)) FROM [TK].dbo.MOCTA  WITH (NOLOCK) WHERE TA011 NOT IN ('Y','y') AND TA006=品號 ) AS '未完成的製令' ");
                sbSql.Append(@"  ,(SELECT CONVERT(INT,ISNULL(MC004,0))  FROM [TK].dbo.BOMMC WHERE MC001=品號) AS 標準批量");
                sbSql.Append(@"  FROM (");
                sbSql.Append(@"  SELECT   TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TC053  AS '客戶' ,TD013 AS '日期',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
                sbSql.Append(@"  ,(CASE WHEN MB004=TD010 THEN (TD008-TD009) ELSE (TD008-TD009)*MD004 END) AS '訂單數量'");
                sbSql.Append(@"  ,MB004 AS '單位'");
                sbSql.Append(@"  ,(TD008-TD009) AS '訂單量'");
                sbSql.Append(@"  ,TD010 AS '訂單單位' ");
                sbSql.Append(@"  ,(CASE WHEN ISNULL(MD002,'')<>'' THEN MD002 ELSE TD010 END ) AS '換算單位'");
                sbSql.Append(@"  ,(CASE WHEN MD003>0 THEN MD003 ELSE 1 END) AS '分子'");
                sbSql.Append(@"  ,(CASE WHEN MD004>0 THEN MD004 ELSE (TD008-TD009) END ) AS '分母'");
                sbSql.Append(@"  FROM [TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.Append(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.Append(@"  WHERE TD004=MB001");
                sbSql.Append(@"  AND TC001=TD001 AND TC002=TD002");
                sbSql.Append(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TC001 IN ({0}) ", TD001.ToString());
                sbSql.Append(@"  AND (TD008-TD009)>0  ");
                sbSql.AppendFormat(@"   {0} ", TC027.ToString());
                //sbSql.Append(@"  AND ( TD004 LIKE '40109916000740%'  ) ");
                sbSql.Append(@"  ) AS TEMP");
                sbSql.Append(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位,單別,單號,序號");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    labelget.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        labelget.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];

                        //建立一個DataGridView的Column物件及其內容
                        DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                        dgvc.Width = 40;
                        dgvc.Name = "選取";

                        this.dataGridView1.Columns.Insert(0, dgvc);

                        dataGridView1.AutoResizeColumns();

                        
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

        public void ADDTOMOCPLANWEEK()
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    MessageBox.Show("號碼 " + ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2] + " 被選取了！");
                }
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
            ADDTOMOCPLANWEEK();
        }
        #endregion


    }
}
