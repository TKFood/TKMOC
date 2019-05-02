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
using FastReport;
using FastReport.Data;
using System.Collections;

namespace TKMOC
{
    public partial class frmPURINVCHECK : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();


        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        ArrayList myAL = new ArrayList();
        string MOCTA001;
        string MOCTA002;
        string MOCTA003;
        string MAXID;

        public frmPURINVCHECK()
        {
            InitializeComponent();
        }
        #region FUNCTION

        public void SEARCHINVMC()
        {
            DateTime SEARCHDATE2 = DateTime.Now;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MC001 AS '品號',MB002 AS '品名',MC002 AS '庫別',MB004 AS '單位',MC004 AS '安全批量',MC005 AS '補貨點'");
                sbSql.AppendFormat(@"  ,ISNULL((SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) ,0) AS '目前庫存'");
                sbSql.AppendFormat(@"  ,ISNULL(((SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) -MC004),0) AS '庫存差異量'");
                sbSql.AppendFormat(@"  ,(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MC001 AND TD007=TD007 AND TD012>='{0}') AS '總採購量'", SEARCHDATE2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,(SELECT TOP 1 ISNULL(TD012,'')+' 預計到貨:'+CONVERT(nvarchar,CONVERT(DECIMAL(16,2),NUM))  FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MC001 AND TD007=TD007 AND TD012>='{0}') AS '最快採購日'", SEARCHDATE2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,(ISNULL((MC004-(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) ),0)-(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MC001 AND TD007=TD007 AND TD012>='{0}') ) AS '需採購量' ", SEARCHDATE2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMC,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE MC001=MB001");
                sbSql.AppendFormat(@"  AND MC002=@MC002 AND MC003='201904制定'");
                sbSql.AppendFormat(@"  ORDER BY ((SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) -MC004),MC001");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                adapter.SelectCommand.Parameters.AddWithValue("@MC002", "20004");

                sqlCmdBuilder = new SqlCommandBuilder(adapter);


                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = ds.Tables["ds"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //ds
                    ds.Tables["dsINVMC"].Rows.Add(row);
                    
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;//調整寬度(標題+儲存格)

                        //dataGridView1.AutoResizeColumns();
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

        public void ADDPURTAB()
        {
            myAL.Clear();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if(Convert.ToDecimal(row.Cells[10].Value.ToString())>0)
                {
                    myAL.Add(Convert.ToDecimal(row.Cells[10].Value.ToString()));
                }
                
            }

            //foreach
            foreach (object num in myAL)
            {
                
            }
        }

        public string GETMAXMOCTA002(string MOCTA001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS ID ");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[PURTA] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE [TA001]='{0}' AND TA003='{1}'", MOCTA001, dateTimePicker1.Value.ToString("yyyyMMdd"));
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
                    return null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        MAXID = SETID(ds1.Tables["ds1"].Rows[0]["ID"].ToString(), dateTimePicker1.Value);
                        return MAXID;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }

        }

        public string SETID(string MAXID, DateTime dt)
        {
            if (MAXID.Equals("00000000000"))
            {
                return dt.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(MAXID.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt.ToString("yyyyMMdd") + temp.ToString();
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHINVMC();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            MOCTA001 = "A311";
            MOCTA002 = GETMAXMOCTA002(MOCTA001);

            ADDPURTAB();

            MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);
           
        }
        private void button3_Click(object sender, EventArgs e)
        {

        }

        #endregion


    }
}
